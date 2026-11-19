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

## Stage 4: icon sets (PR 4)

Refined from spec draft 4 (§3 lines 60–90, especially the `iconSet` schema defaults at line 84; §4 row 4 at line 112; §5.7 lines 175–176; §5.10 lines 199–205; §6.4 lines 261–266; §6.7 lines 282–284; §6.8 line 292; §8; §9 items 4/6/10 at lines 328/330/334; §10 item 4 at line 341; §11 line 350), owner decisions OD-7 (as amended) and OD-8, and the merged code on `main` at `bdc2b217`. Stage 3 merged at `1abea535` (PR #74) and the paged a11y fix at `bdc2b217` (PR #76). The branch is `feature/xlsx-cf-s4`. Every `file:line` below was grepped on `bdc2b217`.

**Goal.** Icon-set rules show an icon inside the host cell where Excel shows it, with Excel's bucket choice, or they show nothing. The scope covers:
- base `<cfRule type="iconSet">`;
- standalone priority-bearing `<x14:cfRule type="iconSet">` (the x14-only sets `3Triangles`, `3Stars` and `5Boxes`, and custom sets);
- custom `<x14:cfIcon>` mappings, including the intentional `NoIcons` bucket.

Icons paint in every view (page, interactive, paged/worker, print, PNG/PDF export) on the resident and paged paths. The text box is inset when `showValue="1"`, text is hidden when `showValue="0"`, and the accessible label is derived from the bucket. None of the Stage 1–3 correctness is lost.

The scope (spec §4 row 4, §10 item 4):
- `cfvo` thresholds of type `num`, `percent` and `percentile`, with `gte`;
- the `<iconSet>` attributes `iconSet`, `reverse`, `showValue` and `percent`;
- all 17 base and 3 x14 built-in families;
- custom `cfIcon` and `NoIcons`;
- icon glyph painting and the horizontal text inset;
- screen-reader icon labels.

`formula` cfvo thresholds, reference `xm:f` thresholds and `activePresent="1"` stay unresolved (`cf-unsupported-rule`) until Stage 5 (spec §3 x14 bullet, §4 fail-closed staging rule). There is no eager glyph library (spec §11 line 350).

**Architecture delta.** Stage 4 adds one visual family. It generalises the Stage 3 seams instead of duplicating them. It adds no second priority mechanism, no new lazy chunk, and no painter outside `src/xlsx/cf/`.
- **Model.**
  - The `iconSet` stub (`parse.ts:1367–1372`) currently ends in `finishUnresolved('cf-unsupported-rule', withDxf(slot))`. It gets a resolved `CfDefinition` arm.
  - Its write mask, `new Set([type])`, becomes the canonical `ICON_SET_WRITES = {'iconSet'}` for resolved and unresolved icon sets.
  - `CfCellState.visual.iconSet?: { name; index }` (`model.ts:64`) is replaced in place by the evaluated descriptor, as Stage 3 did for the bar scalar. A grep of `src/` and `tests/` finds no consumer of the old shape. The only `iconSet` hits outside `parse.ts` are the capability id `priority:cross-visual-iconSet` (`verifiedCapabilities.ts:69,351`) and the manifest.
- **Recognition gap on main.**
  - `ICON_SETS` (`parse.ts:151–155`) lists the 17 base and 3 x14 families but not `NoIcons`. `scanRuleVisual` (`parse.ts:1481–1498`) therefore refuses every `<x14:cfIcon iconSet="NoIcons" iconId="0"/>` as `cf-unsupported-icon`, which contradicts spec §5.10 line 203 ("valid output, not a refusal").
  - Stage 4 admits `NoIcons` only as a `cfIcon` target with `iconId="0"` (S4-B).
  - The limits already fit icon sets: `CF_LIMITS.cfvo: 5` and `customIcons: 5` (`limits.ts:29`), and `cfChildLimit` already bounds `cfvo`/`cfIcon` at 6 with one overflow witness (`parse.ts:2109`). The streamed capture already keeps `cfIcon` (`src/xlsx/cf/capture.ts:146,240`).
- **Dialect.**
  - Main refuses every standalone x14 colour scale (OD-3, `parseColorScaleRule`, `parse.ts:750`).
  - Icon sets cannot do the same, because `3Triangles`/`3Stars`/`5Boxes` and custom sets exist only as standalone x14 rules (spec §3 standalone bullet).
  - Their XML shape comes from an Excel-authored fixture (Lesson 12), not from the generator.
- **Statistics.**
  - Icons reuse `createCfStatistics`/`CfStatisticsSession` (`statistics.ts:397`), `CfRangeStatistics` (`:58`), `percentileOrderStatistic` (`colorScale.ts:41`) and the overflow-safe `lerp` (`colorScale.ts:36`).
  - `resolveScaleThresholds` (`colorScale.ts:282`) accepts only 2 or 3 points and gates on `scale:*` ids. Icons therefore get their own `resolveIconThresholds`, which composes those primitives and contains no sort and no percentile code. It does not widen the scale function a second time.
  - In `service.ts`, `statistical()` (`:17`), `ownSorted` (`:170`), `blockScope` (`:147`, family `'scale' | 'bar'` → `| 'icon'`) and the per-cell branch (`:254`) gain an `iconSet` arm. `estimateCfPrepareBytes` (`:34`) adds `CF_ICON_CELL_BYTES`, the way it adds `CF_BAR_CELL_BYTES` (`:55–57`).
- **Composition.**
  - `visualSlotOf` (`compose.ts:25`) returns `'colorScale' | 'dataBar'`. It gains `'iconSet'`, an independent slot with its own exclusivity and halting gates (S4-I, S4-J).
  - Fill-slot and bar-slot arbitration are untouched.
- **Bytes (Lesson 6).**
  - `CfCellStateWire` (`src/xlsx/view/contracts.ts:396`) gains `icon?`, built from numbers and booleans only.
  - `measureCfCellBytes` (`:481`) adds `'icon'` to its `Pick` and gains a branch. `CF_ICON_CELL_BYTES` is derived from the measurer next to `CF_BAR_CELL_BYTES` (`limits.ts:13`).
  - `pagedSource.ts` gains `hasIcon`/`perCellIcon` and an `icon` omit family, parallel to `hasBar` (`:2000–2005`) and the `families` list (`:2174–2177`).
- **Paint seam (generalised, not duplicated).**
  - `CfPaintBatch` (`src/xlsx/cf/render.ts:23`) has `barAt`/`paintBars`/`hidesTextAt` (`:28–34`).
  - Stage 4 adds `iconAt` and `textInsetAt`. Its own `paintIcons` call sits beside each `paintBars` call: `src/xlsx/render.ts:928` and `src/xlsx/view/renderViewport.ts:305`, in the fill phase before `'content'`.
  - `hidesTextAt` becomes the single text-suppression predicate for both visuals: a painted `showValue="0"` bar or icon.
  - Hosts reuse `residentBarHosts`/`frameBarHosts` (`render.ts:151,186`) unchanged. Their `CfBarHost` (`dataBarPaint.ts:139`) is a plain rect plus fill, so no icon-specific host function is added.
  - Wire states reach the paged main thread through `text.ts` (`CopiedState` at `:106`, `assembleScope` at `:200`, `prepareCfWireScope` at `:514`), which gains `icon`.
- **Text inset (the structurally new piece).**
  - Bars never move text, but an icon does (§5.10 line 204, §6.4 line 264). The viewport path hides text at replay time (`renderViewport.ts:277` strips `body.text`), after layout.
  - An inset, by contrast, changes measurement, the `paintedText` `###` width (`renderViewport.ts:38`) and overflow. It must therefore be decided before layout and must enter the layout key.
  - `layoutCellBody` (`render.ts:1265`, `suppressText` at `:1279`) gains a trailing `inset?`. `CellLayoutHost.layoutBody` (`src/xlsx/view/textLayout.ts:44`) and `createCellLayoutHost` (`renderViewport.ts:29`, its `layoutBody` at `:34`) pass it. `PlacedCell` (`textLayout.ts:88`) gains `textInset?`, and `layoutKey` (`textLayout.ts:288`) includes it.
  - The inset is decided from the placement box with the painter's own predicate (`iconWithheld`), never from a paint-learned map. The paged paint scope's `laidOut` learning (`text.ts:228–234`) cannot drive a pre-layout decision (S4-L, Lesson 11).
- **Glyphs (OD-9).**
  - All glyph data sits behind one interface in a new lazy module, `src/xlsx/cf/iconGlyphs.ts`: `iconGlyph(set, iconId)` returns a glyph or `undefined`.
  - The painter, the wire and the a11y path never see glyph data. The worker does evaluate CF (it loads the lazy chunk; `prepareCfRemoteOnce`, `pagedSource.ts:1772`), but glyph data is static paint-side data in the lazy chunk, so the wire carries ids only and never glyph bytes.
  - OD-9 is decided as (b), original vectors with no Microsoft artwork. That decision shapes Task 28 only.
- **Accessibility.**
  - `describeCfWin` (`describe.ts:65`) gains an `icon` parameter after `bar`, with the label derived from the descriptor's bucket semantics (Lesson 10).
  - The cache key is the existing `descriptorKey` in `SemanticPresentationCache` (`src/components/xlsx/formatPresentation.ts:66,84`). The spec's `cfA11yKey` (§6.7 line 281) is still not a symbol in the code.
- **Invalidation.** It stays whole-sheet, pinned by `tests/xlsx-cf-live.test.tsx:342` and its bar-only twin at `:482`.

**Global Constraints delta.** The Stage 1–3 Global Constraints, binding requirements and common gates apply unchanged. Stage 4 adds or re-binds:
- **Eager budget (Lesson 7).**
  - The eager bundle stays strictly below 38,720 B gzip (`scripts/eager-budget.ts:14`).
  - `src/xlsx/cf/` (including the new `iconSet.ts`, `iconPaint.ts` and `iconGlyphs.ts`) stays in `EAGER_FORBIDDEN`/`PAGED_WORKER_FORBIDDEN`.
  - The only eager additions are the `iconAt`/`paintIcons`/`textInsetAt` call sites, the `inset` parameter, and the `PlacedCell.textInset` key field. Their gzip growth is recorded in `scripts/eager-baseline.json` and passes `bun run gates:core` (+16 B allowance).
  - Glyph data lives only in the lazy `cf` chunk (`vite.config.ts:66`).
- **Exact or withheld (Lesson 2).**
  - Each icon cell gets Excel's bucket, Excel's box and Excel's inset, or no icon and no inset.
  - A ruling default never simplifies a rendering: there is no "icon without inset", no "inset without icon", and no "nearest bucket".
  - Every `[measure]` branch refuses until its named `icon:*` capability is verified. Glyph artwork follows OD-9(b): original vectors, labelled `pixel-approximated`/`*unverified*`, each glyph withheld unless it passes the box and colour check.
- **Evidence (Lesson 4).**
  - `verifiedCapabilities.ts` changes only via `bun scripts/cf-refs/audit.ts --stage 4 --write`, which must be idempotent.
  - Mirroring Ruling S3-WRITE (`audit.ts:648,695`), the Stage 4 export is carried over verbatim unless `--publish-stage4` is passed. Only Task 30 passes it.
  - The stage gate is dual: audit `stageComplete=true` AND conformance green under production evidence, with no injection.
- **Invalidation (Lesson 5).**
  - Any CF generation change that keeps the lease invalidates the whole sheet, icon-only sheets included.
  - An inset change is a layout change. It invalidates through the same whole-sheet path, and the layout cache misses through `cfKey` (`renderViewport.ts:245`) plus `textInset` in `layoutKey`.
- **Bytes and numerics (Lesson 6).**
  - Icon wire bytes come from `measureCfCellBytes`.
  - Glyph bytes are measured from the encoded table, not hand-counted.
  - Threshold arithmetic stays finite near ±1e308, through `lerp` and `safeFraction` (`colorScale.ts:36,431`).
- **Bounded work (Lesson 9).**
  - Icon evaluation, bucket search and glyph decoding spend deterministic work units under `CF_LIMITS.activeUnits` (`limits.ts:33`), never wall time.
  - Statistics over 20k+ cells use `scan: true` budgets (`limits.ts:257`) and share the session cache with scales, bars and `top10`.
  - A fallback attempt, such as the unsorted retry at `service.ts:191`, gets its own fresh budget.
  - `scripts/cf-refs/active-units.ts` is re-run so that `activeUnits` stays at least 5× the largest production call with icons present.
  - CF ranges never enlarge static, print or export extents.
- **CI (Lesson 8).**
  - osacompile and Excel tests use `it.skipIf(!hasOsacompile)` (`tests/xlsx-cf-oracle.test.ts:12,992`).
  - Clock-dependent harnesses pin `process.env.TZ` (conformance `:1012–1031`).

### Stage 4 scope rulings

Each ruling names the spec clause it interprets, its fail-closed default and its cost if wrong. "Unresolved" means the spec §5.7 atomic fail-closed outcome:
- the icon is withheld and no inset is applied;
- an unresolved `stopIfTrue` rule withholds everything after it;
- any other unresolved rule taints the `iconSet` slot for the rules after it.

A default never ships as a measured fact. Only the named `icon:*` capability lifts a ruling, through `hasEvidence` (`evidence.ts:88`).

- **Ruling S4-A: evidence authority for icons (OD-7 as amended).**
  - **Spec tie:** §8 (reference evidence) and §5.10 (every icon outcome is `[measure]`).
  - **No object-model leg.**
    - The settings readback that OD-7 originally relied on is unreachable on Excel 16.114 for bars (`docs/validation/xlsx-cf.md:445`, recipes-stage3 `:33`). The sdef shows the same construction for icons: `icon set format condition` (`X315`, Excel.sdef `:3362–3374`, with `icon criteria`/`icon criterion icon` at `:3355–3368`) is the type of no element or property. It appears only as an enumerator at `:6605/:6614/:6623`.
    - Stage 4 is therefore pixel-first from the start. The one-shot reachability probe `stage4-icon-om-reachability` records what the object model returns. It lifts nothing, and an unexpected success becomes an owner question, never an automatic evidence kind.
  - **Evidence kinds:**

  | Item | Trusted source | Evidence kind |
  |---|---|---|
  | Which glyph a cell shows: bucket choice, `gte`, `reverse`, custom mapping, `NoIcons`, participation, composition, `stopIfTrue` effect | The cell's icon crop, matched against the glyph key (below) | `pixel-classified` |
  | Icon box (left offset, width, height, vertical position), font-size and row-height dependence, clipping, text ink shift (inset), `showValue="0"` canvas outcome | Column/row scan of the `copy picture` bitmap, using the S3-A procedure steps 1–7 (lock check, capture at 1 px/pt, calibration, raw px, per-capability tolerance, discrimination) | `pixel-measured` |
  | Glyph artwork | Original vectors (OD-9(b), owner-decided). They are compared with the glyph-key crops for box and dominant colour only. | `pixel-approximated`, labelled `*unverified*` and never relabelled measured. No Microsoft artwork is shipped. |

  - **Glyph key (`stage4-icon-glyph-key`).** For each of the 75 `(set, iconId)` pairs, one single-cell x14 custom rule maps every bucket to that same pair.
    - The cell therefore shows that glyph whatever its bucket. Each key rule uses `num` cfvo (0, 1, …, k−1) and its cell holds 10, above every threshold, so the single-cell scope needs no statistics and never meets a degenerate `percent` scope (S4-D). The key therefore depends on no threshold semantics.
    - **`percent` attribute:** the key template writes `<iconSet>` in exactly the attribute form Excel's own dialog writes, taken from `excel-authored-iconsets.xlsx` (Lesson 12). If the dialog omits `percent`, the template omits it, and the resaved XML must keep that form. Because every bucket maps to the same icon, no value of `percent` can change the glyph a key cell shows, so S4-F's open question cannot perturb the key. The attribute is pinned in the Task 24 marker test all the same.
    - Key cells and classified cells share one geometry: column width, row height, font, size, vertical alignment, and number format `;;;`. That format hides the value text while the icon still follows the value, so the crop holds only icon ink against a same-format unformatted control.
  - **Classification rule.**
    - A crop is `pixel-classified` as pair P when it equals P's key crop within ±8 per channel at zero offset. It must also differ from every other key crop of a different glyph by at least one pixel of max-channel delta ≥ 64 (`STRONG_DELTA`).
    - Crops identical across pairs (for example the green up arrow of `3Arrows` and of `4Arrows`) form one glyph class. The class, not the pair, is what classification proves.
    - Matching none refuses with `icon-class-none`, and matching two refuses with `icon-class-ambiguous`.
  - **Preconditions:** `icon:pixel-calibration` (S3-A step 3, rerun in the key workbook) and `icon:glyph-key` (all 75 key cells classified into distinct, stable classes) gate every `pixel-classified` and `pixel-measured` icon capability.
  - **Stub composition:** `icon:composition-unresolved-stub` records an unresolved p1 stub ahead of an icon. The viewer withholds either way (§5.7).
  - **Cost if wrong.** If a key is mis-authored, every classification built on it is wrong. The gate is that the key's own resaved XML must round-trip the 75 single-pair rules (Lesson 12). A misread box is off by at most the tolerance, and `cf-compare` still runs.
- **Ruling S4-B: set names, icon ids and unknown mappings.**
  - **Spec tie:** §5.10 line 203, "An unrecognised `iconSet` name or out-of-range `iconId` refuses the rule (`cf-unsupported-icon`)", and §3 line 84.
  - **Default:**
    - The `iconSet` attribute must name one of the 20 families in `ICON_SETS`. An omitted attribute means `3TrafficLights1` only with `icon:settings-schema-defaults`.
    - A `cfIcon` must name one of those families with `0 ≤ iconId < size`, or `NoIcons` with `iconId="0"`.
    - Anything else (an unknown name, `NoIcons` with another id, a non-integer id, or `custom="1"` with a `cfIcon` count different from the bucket count) refuses the whole contribution: `cf-unsupported-icon`, icon slot tainted, never a partial set.
    - An x14-only family (`3Triangles`, `3Stars`, `5Boxes`) on a base `<cfRule>` is unresolved until `stage4-icon-dialect` records Excel's outcome.
  - **Cost if wrong:** third-party files with legal but unusual mappings stay plain.
- **Ruling S4-C: cfvo structure.**
  - **Spec tie:** §3 `CT_Cfvo` and §5.10 line 201 ("`<cfvo>[0]` is the lowest bucket floor").
  - **Count:** exactly k cfvo children for a k-icon family, or for a custom set's bucket count. Any other count → `cf-invalid-rule`, and more than `CF_LIMITS.cfvo` → `cf-resource-limit` before construction.
  - **Missing `type`:** a cfvo without `type` → `cf-invalid-cfvo`, as main already does (`parse.ts:1369`).
  - **Values:** values resolve only as numeric literals under the Stage 1 grammar (`cfvoLiteral`, `parse.ts:952`). A reference, an expression or `type="formula"` → `cf-unsupported-rule` (Stage 5).
  - **`min`/`max`:** these types on an icon cfvo, which Excel's UI does not author, are unresolved until `icon:thresholds-min-max` is measured.
  - **The floor:** the value and type of `cfvo[0]`, and whether values below it draw bucket 0 or no icon, are `[measure]` (`icon:thresholds-floor`). Until then a cell whose value is below a non-trivial floor (anything other than `percent 0`, `percentile 0`, or `num` ≤ the scope minimum) is unresolved.
  - **Cost if wrong:** icon sets with unusual floors stay plain.
- **Ruling S4-D: thresholds and `gte`.**
  - **Spec tie:** §5.10 line 202 (`>=` when `gte` is true, `>` when false `[measure]`), and §9 item 4 (line 328).
  - **Hypothesis:** with resolved thresholds t₁ … t_{k−1}, bucket(v) is the largest i such that v ≥ tᵢ (or v > tᵢ when `gte="0"`), and 0 when there is none.
  - **Threshold resolution:**
    - `percent` p resolves to `lerp(min, max, p/100)`;
    - `percentile` p resolves to `percentileOrderStatistic(sortedAsc, p, variant)`, with the variant chosen by `icon:thresholds-percentile-inc-exc`. It is measured for icons, not inherited from `scale:*`;
    - `num` resolves to the literal.
  - **Gates:** each type needs its own `icon:thresholds-*` capability, and `gte="0"` needs `icon:thresholds-gte-false`.
  - **Layer (as bars split it):** parse-time gates (`iconRequiredCapabilities`, refusing `cf-unsupported-rule`) cover dialect, schema defaults, custom mapping, `NoIcons`, the `percent` attribute and RTL. Evaluation-time gates (`iconEvaluationRequiredCapabilities`, refusing `cf-unmeasured-capability`) cover family order, threshold types, `gte`, the floor, `reverse` and degenerate data.
  - **Ids:** `icon:thresholds-num`, `icon:thresholds-percent`, `icon:thresholds-percentile-inc-exc`, `icon:thresholds-gte-true` (also required for the default `gte`) and `icon:thresholds-mixed` (a list mixing types).
  - **Degenerate data:** non-monotonic thresholds (tᵢ < tᵢ₋₁ after resolution; `icon:degenerate-non-monotonic`), all values equal (`icon:degenerate-equal-values`), a single numeric cell (`icon:degenerate-single-cell`), and `percent`/`percentile` outside 0..100 (`icon:bounds-out-of-domain`) are unresolved until `stage4-icon-degenerate` measures them.
  - **Discrimination (Lesson 3).**
    - `gte` is judged only on cells whose value equals a threshold exactly. The fixture uses binary-exact data and thresholds (`num` 20/40; `percent` 25/50/75 over min 0 and max 100), so the boundary is not lost to rounding. Adjacent buckets also show different glyph classes.
    - Off-boundary cells are designated non-discriminating for `gte`.
  - **Cost if wrong:** boundary cells keep their value with no icon until measured.
- **Ruling S4-E: `reverse` and custom order.**
  - **Spec tie:** §5.10 line 202, "`reverse="1"` reverses the icon assignment order"; §9 item 4.
  - **Family order:** for a built-in family, which icon id bucket b shows (the hypothesis is id b) is `[measure]`. It is enabled only by `icon:families-base` (the 17 base families) or `icon:families-x14` (`3Triangles`, `3Stars`, `5Boxes`), settled by `stage4-icon-families` against the glyph key. Until then every built-in, non-custom icon is unresolved.
  - **Default:**
    - `reverse` maps bucket b to icon (k−1−b) only with `icon:reverse`.
    - `reverse="1"` on a custom set is unresolved until `icon:reverse-custom` is measured, because whether the reversal applies before or after the `cfIcon` mapping is not stated.
  - **Discrimination (Lesson 3):** the middle bucket of an odd-k set maps to itself, so it is non-discriminating for `reverse`. Only the end buckets are designated. A custom set whose end icons are one glyph class is non-discriminating as a whole.
  - **Cost if wrong:** reversed sets stay plain until measured.
- **Ruling S4-F: the `<iconSet percent>` attribute.**
  - **Spec tie:** §3 line 84 ("`percent` (on `<iconSet>`) defaults to `true`"). This attribute is distinct from the cfvo `type="percent"`.
  - **Default:**
    - `percent="1"` or omitted needs `icon:settings-schema-defaults`.
    - `percent="0"` is unresolved until `icon:attr-percent-false` records what Excel does with it, for example whether it reinterprets `percent` cfvo values.
    - The attribute never changes cfvo types in the viewer.
  - **Cost if wrong:** rare third-party files stay plain.
- **Ruling S4-G: participation.**
  - **Spec tie:** §5.10 line 205, "Text, boolean, error and blank cells draw no icon `[measure]`", and §5.1 for date serials.
  - **Per-cell table, as S3-G:**
    - number → evaluated;
    - date serial → evaluated only with `icon:participation-dateserial`, else unresolved, with `numberFormats === 'resolved'` required as at `service.ts:166`;
    - text, boolean, error, `""` and an unstored blank → no icon only with `icon:participation-types`, else unresolved (slot tainted);
    - an error anywhere in a statistic-bearing scope → the whole rule is unresolved unless `icon:participation-error-presence` is verified.
  - **`num`-only rules:** a rule whose thresholds are all `num` needs no statistics. It is not blocked by a scope error, but the erroring cell itself still follows the per-cell table.
  - **Cost if wrong:** lower-priority icons on text cells stay missing until measured.
- **Ruling S4-H: scope statistics.**
  - **Spec tie:** §5.7 multi-range (line 177) and §6.3.
  - **Default:** `blockScope` gains the `'icon'` family. These shapes are each unresolved until their `icon:scope-*` capability is verified, measured for icons and not inherited: disjoint `sqref` (`icon:scope-disjoint`), duplicate blocks (`icon:scope-duplicate-blocks`), hidden row (`icon:scope-hidden-row`), autofilter-hidden row (`icon:scope-autofilter-hidden-row`), merged host (`icon:scope-merged-cell`), and whole column (`icon:scope-whole-column`).
  - **Merged hosts:** a merged host also withholds the paint, as for bars (`src/xlsx/cf/render.ts:412–416`).
  - **Cost if wrong:** multi-range or filtered icon sets stay plain.
- **Ruling S4-I: slot exclusivity and cross-visual composition.**
  - **Spec tie:** §5.7 line 176, "At most one `dataBar` and one `iconSet` render per cell … `[measure]`".
  - **Two icons:** two matching icon rules on one cell withhold both and taint `iconSet`, until `icon:composition-overlapping-icons` is verified. Then the first in priority order paints, as in Ruling S3-I-OVERLAP.
  - **Icon with a bar:** an icon and a bar on one cell paint together only with `icon:composition-with-bar`, which also measures the icon's z-order relative to the bar. Until then the icon is withheld and the bar keeps its own Stage 3 outcome.
  - **Conflicting `showValue`:** if the bar says `showValue="1"` and the icon `"0"` (or the reverse), text suppression is unresolved and the text stays, until `icon:composition-bar-showvalue` is verified.
  - **Fills:** an icon never changes the fill slot. It paints over a scale or DXF fill only with `icon:composition-with-scale` or `icon:composition-dxf-fill`.
  - **`NoIcons`:** whether a `NoIcons` bucket occupies the slot, halting a lower icon rule, is `[measure]` (`icon:composition-noicons-slot`). Until then, a lower matching icon rule under a `NoIcons` win is withheld.
  - **Halting:** a matching DXF `stopIfTrue` ahead of an icon halts it. Until `icon:composition-matching-stopIfTrue-ahead` is verified, the icon is withheld, which is also what halting would show.
  - **Cost if wrong:** icons over filled or barred cells stay plain until measured.
- **Ruling S4-J: an icon rule carrying `stopIfTrue="1"` (OD-1 analogue).**
  - **Spec tie:** §5.7 line 175, which is silent on whether a visual "matches".
  - **Default:** until `icon:composition-icon-stopIfTrue` is verified, the rule is unresolved, so later rules on its cells are withheld (`tainted.add('ALL')`, as the S3-J path does through `halts` at `compose.ts:65–66`).
  - **Lesson 12:** if Excel drops the flag on resave (S3-CLOSE did so for bars), the capability is recorded from the input workbook and stays unverified.
  - **Cost if wrong:** later formatting under such icons is withheld.
- **Ruling S4-K: icon box, placement, clipping and zoom.**
  - **Spec tie:** §5.10 line 204 (the box width plus gap is `[measure]`) and §6.4 line 266 (clip to host, snap per §8, tile seams at 50/100/175 % × DPR 1/2).
  - **Default:** the box rect follows a closed variant set, `CfIconBoxVariant = { size: 'fixed' | 'font' | 'row'; vertical: 'bottom' | 'center' | 'top' | 'cell-align' }`. Exactly one variant is enabled, by `icon:geometry-box` and `icon:geometry-vertical-align` evidence.
  - **Fixtures:** they vary font size (8, 11, 20 pt), row height (12, 30, 60 pt) and vertical alignment, so each pair of hypotheses differs by ≥ 3 px at a designated cell (S3-A step 7).
  - **Edge cases:** a cell narrower or shorter than the box withholds its icon until `icon:geometry-narrow-clip`/`icon:geometry-short-row-clip` is verified.
  - **Constants:** they are pinned in `CF_ICON_RENDER_MODEL` and checked against the sidecars.
  - **Zoom (OD-8 extends to icons):** at other zoom/DPR values the measured point geometry is scaled and snapped to device px, with fidelity labelled `*unverified*`. Buckets stay exact.
  - **Cost if wrong:** sub-pixel drift at non-100 % zoom.
- **Ruling S4-L: text inset and `showValue` on canvas.**
  - **Spec tie:** §5.10 line 204 and §6.4 lines 263–264.
  - **One predicate:** the inset is applied, and text hidden, only on a host whose icon this frame actually paints. The decision is made before layout from the placement box with `iconWithheld(icon, host)`, the painter's own predicate. Every input of that predicate is known before layout:
    - evidence;
    - the host rect;
    - the glyph admission, which S4-R makes at batch preparation, independent of scale;
    - the text width needed by the `icon:inset-overflow` fallback, carried on the host as a typed optional field. `CfBarHost` (`dataBarPaint.ts:139`) widens to `Rect & { fill?: string; textWidth?: number }`, and `CfFrameHostSource.cell` (`render.ts:180–182`, today `{ box, base }`) gains an optional `textWidth`. Paint legs supply it from layout. The paged semantic leg supplies it from the frame's own laid-out record through `CfFrameHostSource.cell` → `frameBarHosts`, and the resident semantic leg through `residentBarHosts` from the static renderer's measurer. An unknown width withholds while `icon:inset-overflow` is unverified.
    - In coarse frames, only the retained active cell has an exact host (`renderViewport.ts:289–306`). It alone gets inset or hiding, and the icon pass paints exactly that host (Lesson 11).
    - The paged paint scope's paint-learned `laidOut` map is not consulted for icons.
  - **The inset:** the inset width comes from `icon:inset-left` on left-aligned numbers in columns wide enough that no `###` appears.
    - The effect on general/right-aligned numbers (`icon:inset-right-aligned`) and centred text (`icon:inset-center`) is measured separately.
    - A right-aligned number in a wide column is non-discriminating for the inset width (Lesson 3).
    - Whether an inset that leaves too little room turns a number into `###`, clips it or overlaps it is `[measure]` (`icon:inset-overflow`). Until then, a cell whose inset text would not fit withholds its icon and keeps its un-inset text.
  - **`NoIcons`:** a `NoIcons` bucket applies no inset and suppresses no text, until `icon:noicons-inset` and `icon:noicons-showvalue` say otherwise (§5.10 line 203 `[measure]`).
  - **`showValue="0"`:** the canvas hides the text only for a painted icon whose `showValue` is `false`, with `icon:showvalue-canvas-suppression`. Whether the icon also moves (for example, centred) when the text is hidden is part of `icon:geometry-box`.
  - **Cost if wrong:** text is shown un-inset with no icon. That is more information, never a wrong value.
- **Ruling S4-M: reading order (RTL).**
  - **Spec tie:** §5.10 line 204 ("or right in RTL `[measure]`").
  - **Default:** on a sheet whose `CfParseMeta.sheetRightToLeft` (`parse.ts:67–73`) is `true` or `undefined`, every icon is unresolved throughout Stage 4, because the viewer does not lay out RTL sheets.
  - **Evidence:** `stage4-icon-rtl` records Excel's placement only (`icon:placement-rtl-sheet` stays unverified, as S3-K did).
  - **Cost if wrong:** icons on RTL sheets stay plain.
- **Ruling S4-N: accessibility, copy and find (OD-2 analogue).**
  - **Spec tie:** §6.7 lines 282–284. The icon label is localised, nothing is added for `NoIcons`, and the `showValue="0"` AX-tree outcome is `[measure]`. That outcome is not observable from the sdef.
  - **Default:**
    - The value text always stays in the accessible name, copy and find. `icon:showvalue-copy-text` records Excel's clipboard text, and a disagreement is a finding, never a viewer change.
    - The descriptor comes from the descriptor's `(bucket, buckets)` semantics, never from paint geometry, glyph data or the glyph class (Lesson 10). The default wording is positional: zh-TW 「圖示 2/3」 and en "Icon 2 of 3".
    - OD-10 (owner-decided 2026-11-14) keeps the positional wording for PR 4. Named labels (for example "green up arrow", from the sdef `XlIcon` names at Excel.sdef `:6393–6446`) are a follow-up and are not built in Stage 4.
    - A cell with both a painted bar and a painted icon announces one bounded label, the bar's then the icon's, joined by the locale joiner, under key `cf:dataBar+iconSet`. If the combined label exceeds `CF_LIMITS.descriptorUnits` (128), only the bar label is announced.
    - The descriptor carries no RGB, thresholds, set names or formulas.
  - **Paged leg:** the paged semantic scope decides with `frameBarHosts` and `iconWithheld`, exactly as S3-A11Y-PAGED does for bars, so the resident, worker and forced-paged legs announce identically.
  - **Cost if wrong:** less informative, never false.
- **Ruling S4-O: dialect and pairing.**
  - **Spec tie:** §3 standalone bullet.
  - **Base:** a base `<iconSet>` needs `icon:dialect-base`.
  - **Standalone x14:** a standalone x14 `<x14:iconSet>` with `@priority` and `<xm:sqref>` needs `icon:dialect-standalone`.
  - **Paired base and x14:** a base rule with an `<x14:id>` companion is unresolved until `stage4-icon-dialect` records whether Excel writes one. Main's `cf-invalid-x14-link` handling for orphans and duplicates is unchanged.
  - **Primary source (Lesson 12):** `tests/fixtures/xlsx-cf/excel-authored-iconsets.xlsx`, written by Excel's own dialog, is the source for the x14 XML. The generator's templates are diffed against it.
  - **Cost if wrong:** x14 icon sets stay plain until their dialect is measured.
- **Ruling S4-P: activation.**
  - **Spec tie:** §3 x14 bullet and §4.
  - **Default:** `activePresent="1"` on an icon set → `cf-unsupported-rule` until Stage 5. The check at `parse.ts:1257` already admits `iconSet` as a visual, so it stays the right shape.
  - **Cost if wrong:** activated icon sets stay plain until Stage 5.
- **Ruling S4-Q: wire form and shared definitions.**
  - **Spec tie:** §6.8, including line 292 (shared definitions such as custom icon descriptors).
  - **Payload:** the per-cell descriptor travels as `CfCellStateWire.icon = { ruleId, set, iconId, bucket, buckets, showValue }`. All fields are numbers and booleans. `set` is an index into the frozen `CF_ICON_SET_NAMES` table, which has 21 entries including `NoIcons`.
  - **No shared table (owner-accepted 2026-11-14):**
    - Spec §6.8 line 292 lists custom icon descriptors among the shared per-sheet definitions. It binds the `RemoteSparklineCache` lifecycle (grant, transfer and adoption, cancellation, eviction, close) to any such definition the main thread holds.
    - In Stage 4 the worker resolves a custom mapping per cell to one `(set, iconId)`. The main thread then needs no per-rule icon data: painting, inset, hiding and the label all read the per-cell descriptor, and the glyph table is static code in the lazy chunk, not transferred data.
    - So no shared icon table exists, and `CfSharedDefinition` (`contracts.ts:409`) is unchanged.
    - The per-cell icon states are not exempt from §6.8. They ride the existing `CfCellStateWire` reply and `RemoteCfCache` (`viewerRemoteCf.ts:147`) adoption accounting, so they inherit that grant, adoption and close lifecycle.
    - Task 27's `the main thread consumes no per-rule icon data` (an acceptance test plus a source grep) proves this.
    - If a later measurement needs per-rule data on the main thread, it must adopt the `RemoteSparklineCache` lifecycle and be priced in grants. This plan adds none.
  - **Pricing:** the payload is priced before construction from `CF_ICON_CELL_BYTES`.
  - **Overflow:** an icon field that exceeds the grant is omitted, with the base cells returned and a `cf-grant-required` diagnostic if it fits. Pending pages yield `pendingSheetIds`, never provisional icons.
  - **Cost if wrong:** reply headroom only.
- **Ruling S4-R: chunking, glyph encoding and charges.**
  - **Spec tie:** §6.1, §7 and §11 line 350.
  - **Placement:** icon code and the glyph table live in the lazy `cf` chunk.
  - **Decoding:** each glyph class a batch uses is decoded once, scale-independently, at batch preparation, charged before allocation to the batch's paint memory (`paintMemoryFor`, `render.ts:450`) and spending work units. Drawing at any zoom/DPR allocates nothing. A refused charge withholds that class's icons, and their insets, for the batch's whole life, never the cell. So no cached inset layout can outlive a refusal.
  - **No-icon sheets:** a sheet without icon rules decodes nothing.
  - **Reporting:** Task 30 reports the `cf` chunk delta against PR 3.
  - **Cost if wrong:** a larger `cf` chunk; the eager bundle is unaffected.

### Planned files and interfaces

| File | Stage 4 responsibility |
|---|---|
| `scripts/cf-refs/cases.ts`, `generate.ts`, `capture.ts`, `verify.ts`, `build.ts`, `audit.ts`, `scripts/cf-compare.ts` | `CF_STAGE4_*`; icon markers; `iconEvidence`; glyph key; OM reachability probe; icon locate/classify; inset measurement; Stage 4 audit and `--publish-stage4`; observed icons |
| `tests/fixtures/xlsx-cf/excel-authored-iconsets.xlsx` (new, orchestrator) | Excel-dialog-authored x14 and custom icon XML (Lesson 12) |
| `src/xlsx/cf/model.ts`, `parse.ts`, `evidence.ts` | Icon definition and descriptor; `NoIcons` admission; both-path icon normalisation |
| `src/xlsx/cf/iconSet.ts` (new), `service.ts`, `compose.ts`, `limits.ts`; `src/xlsx/view/contracts.ts` | Thresholds, buckets, slot composition, measurer branch, `CF_ICON_CELL_BYTES` |
| `src/xlsx/view/pagedSource.ts`, `src/worker/viewerRemoteCf.ts`, `src/xlsx/cf/text.ts` | Icon wire under grants, pending pages, whole-sheet invalidation |
| `src/xlsx/cf/iconGlyphs.ts` (new; the only OD-9-dependent file) | Compact glyph table and decoder |
| `src/xlsx/cf/iconPaint.ts` (new), `src/xlsx/cf/render.ts`, `src/xlsx/render.ts`, `src/xlsx/view/renderViewport.ts`, `src/xlsx/view/textLayout.ts`, `src/xlsx/view/formatWorkspace.ts`, `src/xlsx/cf/describe.ts` | `iconAt`/`paintIcons`/`textInsetAt`, fill-phase painting, inset through layout, `showValue` suppression, descriptor |

**Consumes** (all verified on `bdc2b217`):
- **Model and parse:**
  - `CfDefinition`/`CfCellState`/`CfOutcome`/`CfScaleBound` (`model.ts:35,57,6,15`), `CfParseMeta` (`parse.ts:67`);
  - `ICON_SETS` (`:151`), `scanRuleVisual` (`:1451`), `cfvoLiteral` (`:952`), `barRequiredCapabilities` (`:910`, the pattern to mirror);
  - `parseCfSheet`/`finalizeCfSubtrees` (`:2160,2142`), `cfChildLimit` (`:2107`), `CF_LIMITS.cfvo`/`customIcons` (`limits.ts:29`).
- **Evaluation and composition:** `lerp`/`percentileOrderStatistic`/`safeFraction`/`CfPercentileVariant` (`colorScale.ts:36,41,431,24`), `CfRangeStatistics` (`statistics.ts:58`), `createCfStatistics` (`:397`), `visualSlotOf`/`composeCell`/`composeCellAsync` (`compose.ts:25,289,301`), `createCfWorkBudget` (`limits.ts:260`).
- **Evidence:** `hasEvidence`/`withExactCfEvidence`/`injectCfEvidence` (`evidence.ts:88,28,54`), and the `barLayoutEvidence` pattern (`:156`) for exporting a measured variant.
- **Bytes and worker:** `measureCfCellBytes`/`CfCellStateWire` (`contracts.ts:481,396`), `CF_BAR_CELL_BYTES` (`limits.ts:13`, the pattern), `measurePayloadBytes` (`src/xlsx/view/payloadBytes.ts:62`), `measureCfPayload`/`measureCfWalkedBytes` (`limits.ts:131,148`), `admitReply` (`src/worker/viewerReplyLimit.ts:26`), `RemoteCfCache` (`src/worker/viewerRemoteCf.ts:147`), `prepareCfRemoteOnce` (`pagedSource.ts:1772`).
- **Paint, layout and text:**
  - `CfPaintBatch`/`prepareCfPaint`/`paintMemoryFor` (`cf/render.ts:23,197,450`), `residentBarHosts`/`frameBarHosts`/`CfFrameHostSource` (`:151,186,180`);
  - `barHidesText`/`barWithheld`/`CfBarHost`/`CfBarScale` (`dataBarPaint.ts:113,126,139,148`), `paintDataBars` (`:287`, the pattern);
  - `assembleScope`/`CopiedState`/`prepareCfWireScope` (`text.ts:200,106,514`), `EffectiveFormatScope` (`formatWorkspace.ts:18`);
  - `layoutCellBody`/`paintCellLayout` (`render.ts:1265,1433`), `CellLayoutHost`/`PlacedCell`/`layoutKey` (`textLayout.ts:43,88,288`), `createCellLayoutHost` (`renderViewport.ts:29`), `describeCfWin` (`describe.ts:65`), `SemanticPresentationCache` (`formatPresentation.ts:66`).
- **Oracle harness:**
  - `capture` (`scripts/xlsx-refs/capture.ts:20`), `isDisplayLocked`/`parseScreenLockState` (`scripts/cf-refs/build.ts:88,68`), `captureUnderLockGuard` (`scripts/cf-refs/capture.ts:369`);
  - `decodePngSync`/`computeWorksheetRangeGeometry`/`measureBarGeometry`/`discriminateBar`/`verifyCfCase` (`verify.ts:815,991,1706,1957,3958`), `buildDataBarSettingsProbeScript` (`capture.ts:260`, the pattern);
  - `auditStage3Manifest`/`renderVerifiedCapabilitiesSource` (`audit.ts:500,697`), `CF_STAGE3_CASES`/`CF_STAGE3_NONDISCRIMINATING`/`CF_CASES` (`cases.ts:4542,4635,4694`), `CfCellExpectation.barEvidence` (`cases.ts:136`, the pattern), `stage3ObservedBars` (`cf-compare.ts:197`, the pattern).

**Produces** (every name below is new; the creating task is given):
```ts
// Task 24 — scripts/cf-refs
CfPartMarker.icon?: { dialect: 'base' | 'x14'; set: string; custom?: readonly { set: string; iconId: number }[];
  cfvo: readonly { type: string; val?: string; f?: string; gte?: boolean }[]; attrs: Readonly<Record<string, string>> }
CfCellExpectation.iconEvidence?: { bucket: 'pixel-classified' | 'non-discriminating' | 'unverified';
  geometry: 'pixel-measured' | 'non-discriminating' | 'unverified' }
CfCellExpectation.iconApplied?: boolean
CF_STAGE4_CASES: readonly CfOracleCase[]                                                                   // cases.ts
CF_STAGE4_CAPABILITIES: readonly CfCapabilityEntry[]                                                       // cases.ts
CF_STAGE4_NONDISCRIMINATING: Readonly<Record<string, string>>                                              // cases.ts
CfObservedIcon (new): { ref; cellLeftPx; cellTopPx; cellWidthPx; cellHeightPx;
  icon: { leftPx; topPx; rightPx; bottomPx; glyphClass: string } | null; textInkLeftPx?: number; textInkRightPx?: number }
buildIconOmReachabilityScript(workbookName: string, sheetName: string): string[]                         // capture.ts
classifyIconCrop(png: DecodedPng, crop: Rect, key: ReadonlyMap<string, DecodedPng>, opts: { tolerance: number }):
  { glyphClass: string } | Refusal                                                                         // verify.ts
measureIconGeometry(png: DecodedPng, geom: WorksheetRangeGeometry, refs: readonly string[],
  opts: { controlOf: (ref: string) => string; tolerancePx: number }): ReadonlyMap<string, CfObservedIcon> | Refusal  // verify.ts
CfRefusalCode += 'icon-class-none' | 'icon-class-ambiguous' | 'icon-key-mismatch' | 'icon-edge-unstable' | 'icon-hypothesis-none'
auditStage4Manifest(manifest: CfManifest, options?: { rootDir?: string }): Stage4AuditReport               // audit.ts
RenderVerifiedOptions.publishStage4?: boolean                                                              // audit.ts (--publish-stage4)
stage4ObservedIcons(c: CfOracleCase, rootDir: string): Map<string, CfObservedIcon> | undefined            // cf-compare.ts
// Task 25 — src/xlsx/cf/model.ts, parse.ts
type CfIconSetName = '3Arrows' | … | '5Boxes'          // the 20 ICON_SETS keys, as a literal union
CF_ICON_SET_NAMES: readonly (CfIconSetName | 'NoIcons')[]   // frozen, 21 entries; the wire index table
ICON_SET_WRITES: ReadonlySet<string>                        // = {'iconSet'}
type CfIconThreshold = Readonly<{ bound: { type: 'num' | 'percent' | 'percentile'; value: number } | { type: 'min' } | { type: 'max' }; gte: boolean }>
type CfIconRef = Readonly<{ set: CfIconSetName | 'NoIcons'; iconId: number }>
// CfDefinition gains: { kind: 'iconSet'; set: CfIconSetName; thresholds: readonly CfIconThreshold[]; reverse: boolean;
//   showValue: boolean; percentAttr: boolean; custom?: readonly CfIconRef[]; sheetRightToLeft?: boolean }
type CfIconDescriptor = Readonly<{ ruleId: number; set: CfIconSetName | 'NoIcons'; iconId: number; bucket: number; buckets: number; showValue: boolean }>
// CfCellState.visual.iconSet?: CfIconDescriptor   — replaces { name; index } in place
iconRequiredCapabilities(def: Extract<CfDefinition, { kind: 'iconSet' }>, options?: { defaultedAttributes?: boolean }): Set<string>
// Task 26 — src/xlsx/cf/iconSet.ts, compose.ts, limits.ts, contracts.ts
resolveIconThresholds(def, stats: CfRangeStatistics, options?: { hasEvidence?: (cap: string) => boolean }): CfOutcome<readonly number[]>
iconBucket(value: number, thresholds: readonly number[], gte: readonly boolean[]): number
evaluateIconSet(value: CfLiteral | number, def, thresholds: readonly number[],
  options?: { hasEvidence?: (cap: string) => boolean }): CfOutcome<Omit<CfIconDescriptor, 'ruleId'>>
iconEvaluationRequiredCapabilities(def): Set<string>
// visualSlotOf(rule) widens to 'colorScale' | 'dataBar' | 'iconSet' | undefined
CfCellStateWire.icon?: Readonly<{ ruleId: number; set: number; iconId: number; bucket: number; buckets: number; showValue: boolean }>
CF_ICON_CELL_BYTES: number  // measureCfCellBytes({properties:{}, icon: <full>}) − measureCfCellBytes({properties:{}})
// Task 28 — src/xlsx/cf/iconGlyphs.ts (OD-9(b), owner-decided; the only task the decision shapes)
type CfIconGlyph = Readonly<{ glyphClass: string; widthPt: number; heightPt: number; source: 'original'; path: readonly CfGlyphOp[] }>  // OD-9(b): original vector ops only
type CfGlyphOp (new): a compact move/line/arc/fill opcode tuple with an ARGB fill
iconGlyph(set: CfIconSetName, iconId: number): CfIconGlyph | undefined     // undefined ⇒ the icon is withheld
drawIconGlyph(ctx: CanvasRenderingContext2D, glyph: CfIconGlyph, box: CfIconRect, scale: CfBarScale,
  admitted: CfGlyphCache): boolean                                          // false ⇒ class not admitted, nothing drawn
admitIconGlyphs(classes: ReadonlySet<string>, memory: CfMemory, budget: CfWorkBudget): CfGlyphCache | undefined  // at batch preparation
// Also new in Task 28: CfIconRect (the box type drawIconGlyph takes; Task 29's layoutIcon produces it), CfGlyphCache
CF_ICON_GLYPH_CLASSES: ReadonlyMap<string, string>   // `${set}:${iconId}` → glyph class, copied from the stage4 glyph-key sidecar
// Task 29 — src/xlsx/cf/iconPaint.ts, cf/render.ts, text.ts, layout
// Also new in Task 29: CfIconRenderModel. Elsewhere: Stage4AuditReport (Task 24, audit.ts), CF_ICON_DEFINITION_BYTES (Task 25, limits.ts)
CfBarHost.textWidth?: number                         // widened in place (dataBarPaint.ts:139); residentBarHosts/frameBarHosts pass it through
CfFrameHostSource.cell(...) → { box; base; textWidth?: number }   // widened in place (cf/render.ts:180–182)
CF_ICON_RENDER_MODEL: CfIconRenderModel   // box variant, offsets, gap, clip rules — values copied from stage4 sidecars
type CfIconBoxVariant = Readonly<{ size: 'fixed' | 'font' | 'row'; vertical: 'bottom' | 'center' | 'top' | 'cell-align' }>
layoutIcon(icon: CfIconDescriptor, host: CfBarHost, font: { sizePt: number; vertical?: string }, scale: CfBarScale,
  model: CfIconRenderModel): { box: CfIconRect; inset: { left: number; right: number } } | undefined
iconWithheld(icon: CfIconDescriptor, host: CfBarHost | undefined, check?: (cap: string) => boolean): boolean
CfPaintBatch.iconAt(row: number, col: number): CfIconDescriptor | undefined
CfPaintBatch.paintIcons(ctx: CanvasRenderingContext2D, cellRect: (row: number, col: number) => CfBarHost | undefined,
  visible: { r0: number; c0: number; r1: number; c1: number }, scale: CfBarScale): boolean
CfPaintBatch.textInsetAt(row: number, col: number, host: CfBarHost, scale: CfBarScale): { left: number; right: number } | undefined
EffectiveFormatScope.iconAt? / paintIcons? / textInsetAt?          // same signatures, optional
layoutCellBody(..., suppressText?: boolean, inset?: { left: number; right: number })   // new optional trailing parameter
CellLayoutHost.layoutBody(..., allocation?, inset?)  ;  PlacedCell.textInset?: { left: number; right: number }
describeCfWin(..., bar?, icon?: Readonly<{ bucket: number; buckets: number }>)          // new optional trailing parameter
```

### Execution, assignments and review

Implementers are Claude subagents: Sonnet for mechanical tasks, Opus for judgment-heavy ones. opencode, a different model family, reviews every task. Before dispatch, record the actual model of each seat. Never relabel a same-family review as independent. Implementers do not run Excel: the orchestrator executes every host-Excel step (AppleScript and System Events) while the owner keeps the display unlocked. Reports and reviews go in `.superpowers/sdd/<plan>/task-N-report.md` / `task-N-review.md`. After every agent, check `git status` and `git diff --stat` for test files.

| Task | Implementer | Reviewer |
|---|---|---|
| 24 Stage 4 Excel oracle: glyph key, icon cases, classification, conformance | Claude Opus (orchestrator executes host Excel) | opencode |
| 25 Icon model, resident and streamed parsing, evidence gates | Claude Sonnet | opencode |
| 26 Thresholds, buckets, statistics reuse, slot composition, measurer | Claude Opus | opencode |
| 27 Icon wire, worker grant, pending pages, whole-sheet invalidation | Claude Sonnet | opencode |
| 28 Original vector glyph table and decoder (OD-9(b)) | Claude Opus | opencode |
| 29 All-view icon painting, text inset, `showValue`, accessibility, copy/find | Claude Opus | opencode |
| 30 Stage 4 validation, resource audit and PR evidence | Claude Sonnet | opencode |

Task 28 depends on OD-9 and on the Task 24 glyph-key sidecar, but nothing else depends on Task 28 except the paint step in Task 29. OD-9 is decided (b), so Task 28 waits only on the Task 24 glyph-key capture, while Tasks 25–27 proceed in parallel. Until Task 28 lands, `iconGlyph` returns `undefined`, so every icon is withheld and no inset is applied.

### Review Focus (Stage 4)

1. A `gte` capability verified from cells off the threshold, a `reverse` capability verified from the middle bucket, or a custom set whose end icons are one glyph class. Each is non-discriminating and must withhold (S4-D, S4-E, Task 24).
2. A classification that matches two key crops, or a key whose resaved XML lost a single-pair rule. Each must refuse (`icon-class-ambiguous`, `icon-key-mismatch`), never pick one (S4-A, Task 24).
3. An inset applied where no icon paints, or an icon painted with no inset. That covers coarse frames, a withheld glyph, a refused glyph charge, `NoIcons`, and a narrow cell. The inset and the paint come from one predicate on one host (S4-L, Tasks 29, 30).
4. A layout-cache hit that keeps un-inset text after an icon starts painting, or inset text after it stops. `layoutKey` must carry `textInset`, and the `###` width must use the inset width (S4-L, Task 29).
5. `NoIcons` refused at parse (the main-branch gap), a `NoIcons` bucket that insets or hides text without evidence, or an unknown `cfIcon` that refuses only its bucket instead of the whole contribution (S4-B, Task 25).
6. Two icons on one cell; an icon with a bar and conflicting `showValue`; an icon under a matching DXF `stopIfTrue`; an icon carrying `stopIfTrue="1"`; a lower icon under a `NoIcons` win. Each withholds as ruled, and the Stage 3 bar outcome is unchanged (S4-I, S4-J, Task 26).
7. Thresholds at ±1e308, `percent` over an overflowing span, and equal or non-monotonic thresholds. Results are finite or unresolved, with no NaN bucket (S4-D, Lesson 6, Task 26).
8. A value edit that moves `max` across a threshold: every icon in the scope re-buckets in the same generation, and every tile of an icon-only sheet invalidates (Lesson 5, Task 27).
9. An icon payload at grant − 1: base cells and bars return, the `icon` field is omitted, and the bytes match the walker exactly (S4-Q, Task 27).
10. The resident, worker and forced-paged legs announce the same `Icon b of k` label, derived from the bucket and not the glyph. Copy and find keep the value under `showValue="0"`, and a no-icon sheet decodes no glyph and does zero icon work (S4-N, S4-R, Tasks 29, 30).

### Task 24: Stage 4 Excel oracle: glyph key, icon cases, classification, conformance (Claude Opus; review opencode)

**Size:** kept as one task. The harness, the cases and the withheld-only conformance guard are reviewed together, so no case lands without its guard (Lesson 1), and the dual gate already lives in Task 30.

**Dependencies:**
- Stage 3 merged (PR #74), and the paged a11y fix merged (PR #76).
- Host Excel runs only through the orchestrator (AppleScript and System Events, per the automate-Excel-recipes practice) while the owner keeps the screen unlocked.
- Preparing scripts is not evidence. Every case without captures stays `recipe-pending`, and refused.

**Files:**
- Modify `scripts/cf-refs/cases.ts`:
  - add `CF_STAGE4_CASES`, `CF_STAGE4_CAPABILITIES` and `CF_STAGE4_NONDISCRIMINATING` beside the Stage 3 exports (no separate file, which would create the same cycle as before);
  - add `CfPartMarker.icon`, `CfCellExpectation.iconEvidence` and `iconApplied`;
  - extend `CF_CASES`/`CF_CAPABILITIES`.
- Modify the other oracle scripts:
  - `generate.ts`: icon templates diffed against the Excel-authored fixture;
  - `capture.ts`: `buildIconOmReachabilityScript`, captures through `captureUnderLockGuard`;
  - `verify.ts`: `classifyIconCrop`, `measureIconGeometry` and the icon refusal codes;
  - `audit.ts`: the guard at `:1676` accepts 4, `auditStage4Manifest` follows `:500`, and `--publish-stage4` mirrors S3-WRITE (`:648,695,1693`);
  - `scripts/cf-compare.ts`: `stage4ObservedIcons` beside `stage3ObservedBars` (`:197`).
- Modify the tests and fixtures:
  - `tests/xlsx-cf-oracle.test.ts`, `tests/xlsx-cf-oracle-conformance.test.ts`;
  - `tests/fixtures/xlsx-cf/manifest.json`, plus the case `.xlsx` files and `refs/*`;
  - create `tests/fixtures/xlsx-cf/excel-authored-iconsets.xlsx` (orchestrator).
- Create `docs/validation/xlsx-cf-recipes-stage4.md`. Update `docs/validation/xlsx-cf.md`.

**Interfaces:**
- **Consumes:** `CfOracleCase`, `CfPartMarker`, `CfCellExpectation`, `verifyCfCase`, `capture`, `captureUnderLockGuard`, `decodePngSync`, `computeWorksheetRangeGeometry`, `measureBarGeometry`/`discriminateBar` (patterns), `auditStage3Manifest` (pattern), `renderVerifiedCapabilitiesSource`, and the resave receipt. `cf-compare.ts:623` filters `c.stage > stage`, so `--stage 4` includes Stages 1–3.
- **Produces:** the Task 24 block of **Produces**.

The capability prefix is `icon:*`, one capability per variant. Each capability's `failClosedDisposition` cites the S4 ruling it lifts. Each `iconEvidence` value has defined semantics:
- `bucket: 'pixel-classified'`: the S4-A classification rule matched exactly one glyph class, and the designated cells discriminate;
- `geometry: 'pixel-measured'`: located after calibration and the lock checks, stable on three scanlines, and discriminating;
- `non-discriminating`: measured, but it fails the rule;
- `unverified`: anything else.

**Required Stage 4 oracle matrix.** Every row is `[measure]`. No expected bucket, px or glyph is written by the plan or by the code's authors. The spec §8 rules apply: a same-content unformatted control block, positive and negative controls, a repair-free open, and retained markers diffed against Excel's resaved XML (Lesson 12).

| Case ID | Input / question to settle | Capabilities |
|---|---|---|
| `stage4-icon-om-reachability` | Is `icon set format condition` reachable? Read `format condition 1 of range` and its `properties`, `icon criteria`, `reverse icon set order` and `show icon only`. Record only. | — (lifts nothing) |
| `stage4-icon-authored` | The orchestrator authors through Excel's own dialog (System Events): `3Stars`, `3Triangles`, `5Boxes`, a custom set with one `NoIcons` bucket, `reverse`, "show icon only", and a `gte` toggle. The output is `excel-authored-iconsets.xlsx`. Its XML becomes the generator's template source. | `icon:dialect-base`, `icon:dialect-standalone` |
| `stage4-icon-glyph-key` | 75 single-cell custom rules, every bucket mapped to one `(set, iconId)`. Uniform geometry, `;;;` number format, a same-format control, calibration columns 10/30/60 characters wide, and one solid DXF cell. | `icon:pixel-calibration`, `icon:glyph-key`, one `icon:glyph-<set>-<id>` per pair (artwork, per OD-9) |
| `stage4-icon-families` | Each of the 20 built-in families over k values far from every threshold (bucket order only, default thresholds). This settles which icon id each bucket shows. | `icon:families-base`, `icon:families-x14`, `icon:settings-schema-defaults` |
| `stage4-icon-thresholds` | `num` 20/40 and `percent` 25/50/75 over {0..100}, binary-exact. Cells exactly at each threshold under `gte` 1 and `gte` 0; the off-boundary cells are designated non-discriminating. `percent` over {10..110}. `percentile` 33/67 over {1,2,3,4,10}, where inc ≠ exc at a designated cell. Mixed `num`+`percent`. | `icon:thresholds-num`, `icon:thresholds-percent`, `icon:thresholds-percentile-inc-exc`, `icon:thresholds-gte-true`, `icon:thresholds-gte-false`, `icon:thresholds-mixed` |
| `stage4-icon-degenerate` | All values equal; a single numeric cell; non-monotonic `num` 60/30; `percent` −10 and 150; `cfvo[0]` `num` 50 with values below it; `min`/`max` cfvo types; `<iconSet percent="0">`. | `icon:thresholds-floor`, `icon:thresholds-min-max`, `icon:degenerate-equal-values`, `icon:degenerate-single-cell`, `icon:degenerate-non-monotonic`, `icon:bounds-out-of-domain`, `icon:attr-percent-false` |
| `stage4-icon-reverse-custom` | `reverse` on 3-, 4- and 5-icon sets (end buckets designated). Custom mappings across families. Custom + `reverse`. `NoIcons` in the first, middle and last bucket. | `icon:reverse`, `icon:reverse-custom`, `icon:custom-mapping`, `icon:custom-noicons` |
| `stage4-icon-geometry` | Font sizes 8/11/20 pt × row heights 12/30/60 pt × vertical alignment top/center/bottom. A 2-character column. A row shorter than the box. Column widths that put predicted box edges at .25/.5/.75 pt. | `icon:geometry-box`, `icon:geometry-vertical-align`, `icon:geometry-snapping`, `icon:geometry-narrow-clip`, `icon:geometry-short-row-clip` |
| `stage4-icon-inset` | Left-aligned numbers in wide columns: the text ink shift against the control. General (right-aligned) and centred numbers. A column where number plus icon do not fit. `NoIcons` buckets with `showValue` 1 and 0. | `icon:inset-left`, `icon:inset-right-aligned`, `icon:inset-center`, `icon:inset-overflow`, `icon:noicons-inset`, `icon:noicons-showvalue` |
| `stage4-icon-showvalue` | `showValue` 0/1 on canvas. Icon position with the text hidden. Copy text through `copy range`, then `the clipboard as text`. The AX tree is recorded as not observable. | `icon:showvalue-canvas-suppression`, `icon:showvalue-copy-text` |
| `stage4-icon-participation` | A number, a date serial, `TRUE`, text `"5"`, `#N/A`, `""`, an unstored blank, and a formula with a cached number. A percent variant with one error present, and a `num`-only variant with one error present. | `icon:participation-types`, `icon:participation-dateserial`, `icon:participation-error-presence` |
| `stage4-icon-scope` | Disjoint `B2:B4 D2:D4`; duplicate blocks; a hidden row and an autofilter-hidden row each holding the extreme value; a merged `B8:C9` with non-equal values; a whole column `F:F` with no hidden rows. | `icon:scope-disjoint`, `icon:scope-duplicate-blocks`, `icon:scope-hidden-row`, `icon:scope-autofilter-hidden-row`, `icon:scope-merged-cell`, `icon:scope-whole-column` |
| `stage4-icon-composition` | Two overlapping icon rules; an icon with a data bar (both priority orders, and conflicting `showValue`); an icon over a scale fill and over a DXF fill; a matching and a non-matching DXF `stopIfTrue` ahead; an icon with `stopIfTrue="1"` ahead of a DXF fill; a `NoIcons` win over a lower icon rule; an unresolved `expression` stub at p1. | `icon:composition-overlapping-icons`, `icon:composition-with-bar`, `icon:composition-bar-showvalue`, `icon:composition-with-scale`, `icon:composition-dxf-fill`, `icon:composition-matching-stopIfTrue-ahead`, `icon:composition-icon-stopIfTrue`, `icon:composition-noicons-slot`, `icon:composition-unresolved-stub` |
| `stage4-icon-rtl` | An icon set on a `rightToLeft="1"` sheet. Placement and inset recorded only (S4-M). | `icon:placement-rtl-sheet` (stays unverified) |
| `stage4-icon-unresolved-shapes` | Excel's open/repair/strip outcome for: an unknown `iconSet`; `iconId` out of range; `NoIcons` `iconId="1"`; a custom set with fewer `cfIcon` than buckets; a base rule naming `3Stars`; `type="formula"`; `xm:f` `$A$1`; `activePresent="1"`; 2 and 6 cfvo; a cfvo without `type`. Lifts no capability; the viewer refuses either way. | — |

Expected non-discriminating candidates, which the rule confirms or clears after capture:
- `icon:thresholds-gte-*` on off-boundary cells;
- `icon:reverse` on the middle bucket;
- `icon:inset-left` on right-aligned numbers;
- `icon:geometry-snapping` if no edge lands at a fractional point;
- `icon:composition-noicons-slot` if the lower rule's icon would be withheld anyway.

- [x] **Conformance first (Lesson 1).** Add `stage 4 icon cells are withheld until their evidence lands` to `tests/xlsx-cf-oracle-conformance.test.ts`.
  - Every Stage 4 cell must have `CfPrepared.at(...)` either not `ready` or with `.value.visual?.iconSet === undefined`, under production evidence. The paint-side `textInsetAt` assertion joins in Task 29, which creates it.
  - The existing Stage 1–3 assertions (`:496`, `:557–560`, the 15-case Stage 3 guard at `:645–646`) stay byte-unchanged.
  - Icon FINDINGS are withheld-only, extending the guard at `:996`.
- [x] Write `stage 4 cases join the manifest and stage 1–3 cases are unchanged`. Manifest ids must equal the combined `CF_CASES`, and the Stage 1–3 capability ids must be unchanged. Run RED, then GREEN.
- [x] Write `icon marker pins set, cfvo, gte, custom mapping and dialect exactly`. The verifier refuses each of these mutations:
  - `gte` true → false on `cfvo[2]`;
  - reordered cfvo;
  - `cfIcon[1]` `iconId` 2 → 1;
  - `NoIcons` → `3Arrows:0`;
  - `reverse` flipped;
  - `showValue` flipped;
  - `xm:f` `"20"` → `"2"`;
  - a standalone x14 rule moved to the base dialect;
  - the glyph key's `<iconSet percent>` attribute form changed from the Excel-authored form.

  The verifier compares semantics, not raw strings.
- [x] Write `glyph key refuses a mismatched or ambiguous crop`, on synthetic PNGs.
  - A 3×3 key of distinct crops classifies a pasted copy of crop 2 as class 2.
  - A crop differing from every key by a strong pixel → `icon-class-none`.
  - Two key entries identical to the test crop with different glyph classes → `icon-class-ambiguous`.
  - A key cell whose resaved XML lost its rule → `icon-key-mismatch`.
- [x] Write `icon edges and text ink are located by scan within ±1 px`. A 16×16 icon at px 3–18 gives left 3 and right 18. A text ink start at 24 against a control at 4 gives a shift of 20. Scanlines disagreeing by 2 px → `icon-edge-unstable`.
- [x] Write `non-discriminating icon variant keeps its gate closed (Lesson 3)`. All of these leave `hasEvidence` returning `false`:
  - a `gte` capability whose designated cells contain no exact-boundary value;
  - a `reverse` capability judged only on the middle bucket;
  - two box hypotheses 2 px apart.

  At 3 px apart, with exactly one matching hypothesis, the capability is verified.
- [x] Write `audit --stage 4 --write is idempotent (Lesson 4)`. Two runs produce a byte-identical `verifiedCapabilities.ts`. Without `--publish-stage4` the Stage 4 export is carried over verbatim (absent at first), and the Stage 1–3 sets are unchanged.
- [x] Extend the osacompile test with `buildIconOmReachabilityScript` under `it.skipIf(!hasOsacompile)` (Lesson 8). The participation and clock-sensitive cases reuse the TZ pin pattern.
- [ ] **Orchestrator:**
  1. Run `stage4-icon-om-reachability` and record the result.
  2. Author `stage4-icon-authored` through Excel's dialog, save it, and diff its XML into the generator templates (Lesson 12).
  3. Run `bun scripts/cf-refs/build.ts --stage 4 --case stage4-icon-glyph-key` and `capture.ts --stage 4 --case stage4-icon-glyph-key`. If calibration or the key fails, stop and record it, because every classification stays unverified.
  4. Run the remaining cases (`--refresh-clock-date`, `--verify-only`). Close owned workbooks between cases, and never manufacture pixels.
- [x] Inventory the corpus files that contain `iconSet` (ZIP XML; skip `~$`) into `docs/validation/xlsx-cf.md` §4: dialects, families, custom/`NoIcons`, cfvo types, `gte`, and `reverse`. Record every Stage 4 `[measure]` row with its ruling and status.
- [ ] Run `npx vitest run tests/xlsx-cf-oracle.test.ts tests/xlsx-cf-oracle-conformance.test.ts`, `npx tsc --noEmit -p tsconfig.scripts.json`, `npx tsc --noEmit -p .` and `git diff --check`. Commit (`test(xlsx-cf): add Stage 4 icon-set oracle cases`), then report and review. Pending variants do not block Task 25; their branches stay disabled.

### Task 25: Icon model, resident and streamed parsing, evidence gates (Claude Sonnet; review opencode)

**Files:**
- Modify `src/xlsx/cf/model.ts`, `parse.ts` and `evidence.ts`.
- Modify `tests/xlsx-cf-parse.test.ts` and `tests/xlsx-cf-paged-parse.test.ts`. Create `tests/xlsx-cf-icon-parse.test.ts`.
- `src/xlsx/cf/capture.ts` changes only if the deep-equality test shows streamed `cfIcon` or `x14:iconSet` attributes are lost.

**Interfaces:**
- **Consumes:** `parseCfSheet`, `finalizeCfSubtrees`, `scanRuleVisual`, `ICON_SETS`, `cfvoLiteral`, `barRequiredCapabilities` (pattern), `CF_LIMITS.cfvo`/`customIcons`, `CfParseMeta.sheetRightToLeft`, `hasEvidence`, and the Task 24 `icon:settings-*`/`icon:dialect-*` ids.
- **Produces:** the Task 25 block of **Produces**. The old `visual.iconSet` is replaced in place.

- [ ] Write `canonical icon write mask is ['iconSet'] for resolved and unresolved icon sets`. A resolved set, an unresolved set and a set carrying a DXF each carry exactly `ICON_SET_WRITES` plus any DXF writes. Broken links keep `ALL` (`invalidLinkRule`, `parse.ts:1617`).
- [ ] Write `base 3Arrows resolves with schema defaults`. The input is `<iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet>`.
  - With `icon:settings-schema-defaults` and `icon:dialect-base` injected (the parse layer), it yields `set '3Arrows'`, thresholds `[{percent 0, gte true}, {percent 33, gte true}, {percent 67, gte true}]`, `reverse false`, `showValue true`, `percentAttr true`, and no `custom`.
  - With `withExactCfEvidence([])` it yields `cf-unsupported-rule`, writes `{iconSet}`, and one deduplicated diagnostic.
  - An omitted `iconSet` attribute yields `set '3TrafficLights1'` (spec line 84).
- [ ] Write `standalone x14 custom set with NoIcons resolves`. The input is `<x14:iconSet iconSet="3Arrows" custom="1">` with three `<x14:cfvo type="num"><xm:f>…</xm:f></x14:cfvo>` (0, 20, 40, the second with `gte="0"`) and `cfIcon`s `3Flags:0`, `NoIcons:0` and `3Stars:2`, at `@priority 3`.
  - Under evidence, `custom` = `[{set '3Flags', iconId 0}, {set 'NoIcons', iconId 0}, {set '3Stars', iconId 2}]`, the gte list is `[true, false, true]`, and the priority is 3.
  - This is the main-branch gap: today the `NoIcons` entry refuses as `cf-unsupported-icon` (`parse.ts:1490–1492`).
- [ ] Write `unknown icon mappings refuse the whole contribution (S4-B)`. Each of these → `cf-unsupported-icon` with the slot tainted and no partial set:
  - `iconSet="7Moons"`;
  - `cfIcon iconSet="3Arrows" iconId="3"`;
  - `NoIcons iconId="1"`;
  - `iconId="1.5"`;
  - `custom="1"` with 2 `cfIcon`s on a 3-icon set.

  The diagnostics cap stays at 64.
- [ ] Write `icon cfvo structure refusals`, one row per S4-C/O/P ruling:
  - 2 cfvo on a 3-icon set → `cf-invalid-rule`;
  - 6 cfvo → `cf-resource-limit` before construction;
  - missing `type` → `cf-invalid-cfvo`;
  - `formula` cfvo, `xm:f` `$A$1` and `activePresent="1"` → `cf-unsupported-rule`;
  - `min`/`max` cfvo → unresolved without `icon:thresholds-min-max`;
  - base `3Stars` → unresolved without the dialect evidence;
  - `<iconSet percent="0">` → unresolved (S4-F);
  - `activePresent` on `top10` keeps `cf-invalid-rule`.
- [ ] Write `icon sets on an RTL sheet are unresolved on both paths (S4-M)`. A `<sheetView rightToLeft="1">` sheet, and a sheet with a malformed flag, each yield unresolved on the resident and streamed paths. An LTR sheet resolves under evidence.
- [ ] Write `iconRequiredCapabilities is never empty and names each feature`.
  - A base default set needs exactly `{icon:dialect-base, icon:settings-schema-defaults}`.
  - A standalone x14 set needs `icon:dialect-standalone`.
  - A custom set adds `icon:custom-mapping`, plus `icon:custom-noicons` when it maps a `NoIcons` bucket.
  - `<iconSet percent="0">` adds `icon:attr-percent-false`.
  - `min`/`max` cfvo add `icon:thresholds-min-max`.
  - No evaluation-layer id (`icon:thresholds-*`, `icon:families-*`, `icon:reverse`) appears in the parse set.
- [ ] Write `icon definition bytes are measurer-derived (Lesson 6)`. `CF_ICON_DEFINITION_BYTES` equals `measurePayloadBytes` of the largest icon definition (5 thresholds, 5 custom refs, longest names). The parse reservation uses it beside `CF_BAR_DEFINITION_BYTES` (`parse.ts:1174,1574`).
- [ ] Extend `resident and streamed CF are deep equal at every chunk boundary` (`tests/xlsx-cf-paged-parse.test.ts:32`) with a base set, a standalone x14 custom set carrying `NoIcons`, and `3Stars`, with a chunk boundary inside `<x14:cfIcon>` and inside `<xm:f>`.
- [ ] Run `npx vitest run tests/xlsx-cf-icon-parse.test.ts tests/xlsx-cf-parse.test.ts tests/xlsx-cf-paged-parse.test.ts tests/xlsx-cf-bar-parse.test.ts tests/xlsx-cf-scale-parse.test.ts` to record RED, implement, then run GREEN. Then run `npx tsc --noEmit -p .`, `bun scripts/check-eager.ts` and `bun scripts/eager-budget.ts`. Commit (`feat(xlsx-cf): parse icon-set rules behind evidence gates`), then report and review.

### Task 26: Thresholds, buckets, statistics reuse, slot composition and measurer (Claude Opus; review opencode)

**Files:**
- Create `src/xlsx/cf/iconSet.ts`, `tests/xlsx-cf-icon-eval.test.ts` and `tests/xlsx-cf-icon-compose.test.ts`.
- Modify:
  - `src/xlsx/cf/service.ts`: `statistical`, `ownSorted`, `blockScope` with the `'icon'` family, the per-cell branch, and `estimateCfPrepareBytes` via `CF_ICON_CELL_BYTES`;
  - `compose.ts`;
  - `limits.ts`;
  - `src/xlsx/view/contracts.ts`: `CfCellStateWire.icon`, plus the `measureCfCellBytes` `Pick` and branch.
- Extend `tests/xlsx-cf-worker.test.ts`, `tests/xlsx-cf-lazy.test.ts`, `tests/xlsx-cf-statistics.test.ts` and `tests/xlsx-cf-active-units.test.ts`.

**Interfaces:**
- **Consumes:** the Task 25 definition; `lerp`, `percentileOrderStatistic`, `safeFraction`, `CfRangeStatistics`, `createCfStatistics`, `composeCellAsync`, `visualSlotOf`, `measurePayloadBytes`, `createCfWorkBudget`, and `hasEvidence`.
- **Produces:** the Task 26 block of **Produces**.
- **Constraints:** the evaluator stays pure: no DOM, no clock, no glyph data. Every bucket expectation below is a hypothesis run under injected evidence through `withExactCfEvidence`, never an Excel fact; the facts come from Task 24 sidecars.

- [ ] Write `iconEvaluationRequiredCapabilities names each feature`. A default `3Arrows` percent set needs `{icon:families-base, icon:thresholds-percent}`. `3Stars` needs `icon:families-x14`. `gte="0"` adds `icon:thresholds-gte-false`, `reverse` adds `icon:reverse` (and `icon:reverse-custom` on a custom set), and a percentile threshold adds `icon:thresholds-percentile-inc-exc`. A custom set needs no `icon:families-*` id.
- [ ] Write `no threshold evidence ⇒ unresolved`. With no evidence, every icon evaluates to `{ kind: 'unresolved', reason: 'cf-unmeasured-capability' }`, and `CfCellState.visual` carries no `iconSet`.
- [ ] Write `bucket boundaries under the S4-D hypothesis`. Over {0..100} with `num` 20/40 (`gte` true, true):
  - v=19 → bucket 0, v=20 → 1, v=39 → 1, v=40 → 2, v=100 → 2.
  - With `cfvo[1].gte=false`: v=20 → bucket 0.
  - `iconBucket(20, [0, 20, 40], [true, false, true])` returns 0.
  - `iconBucket(40, [0, 20, 40], [true, true, false])` returns 1.
- [ ] Write `percent and percentile thresholds reuse the shared primitives`.
  - `percent` 25/50/75 over {0..100} resolves to `[0, 25, 50, 75]` exactly. Over {10..110} it resolves to `[10, 35, 60, 85]`.
  - `percentile` 33/67 over {1,2,3,4,10} equals `percentileOrderStatistic` under the injected variant ('inc' → `[1, 2.32, 3.68]` within 1e-12 for the hypothesis run).
  - `iconSet.ts` contains no `sort(` and no percentile implementation (source-text assertion).
- [ ] Write `reverse and custom mapping`.
  - A 3-icon `reverse` maps buckets 0/1/2 to icon ids 2/1/0.
  - A custom mapping returns its `(set, iconId)` per bucket.
  - `NoIcons` returns `{ set: 'NoIcons', iconId: 0 }`.
  - Custom + `reverse` without `icon:reverse-custom` → unresolved.
  - `bucket` and `buckets` are always the semantic bucket (pre-reverse) and k.
- [ ] Write `icon evaluation stays finite at ±1e308 (Lesson 6)`.
  - `percent` 50 over a scope of {−1e308, 1e308} resolves to a finite threshold through `lerp`.
  - Every field of every result is `Number.isFinite`.
  - Equal values, a single numeric cell and non-monotonic thresholds are unresolved, with nothing NaN.
- [ ] Write `icons share the statistics session`.
  - A percentile icon set, a min/max scale, a percent bar and a `top10` rule on one 50-row `A:A` scope visit the source at most 50 times, and a warm hit does zero loops.
  - A `num`-only icon set requests no statistics.
- [ ] Write `icon percentile copy reserves before allocation, and the fallback gets a fresh budget (Lesson 9)`. On 100,000 numeric cells at `requiredBytes − 1`:
  - the result is `cf-resource-limit`, while a `num`-only icon set on the same scope stays ready;
  - the over-cap refusal is cached per (scope, generation) with one diagnostic;
  - an upgraded-sorted refusal retries unsorted under a new `createCfWorkBudget({ scan: true })` (spy: two distinct budgets).
- [ ] Write `icon participation and scope gates`. The S4-G table is applied row by row. A text cell taints the icon slot for a lower icon without evidence. An error makes a percent rule unresolved but not a `num`-only rule. Each S4-H shape is unresolved until its `icon:scope-*` evidence is injected, and `bar:scope-*` evidence never lifts it.
- [ ] Write `icon slot composition` in `tests/xlsx-cf-icon-compose.test.ts`:
  - two matching icons → no icon, with `unresolvedWrites` containing `iconSet`;
  - with `icon:composition-overlapping-icons`, the first in priority order wins;
  - icon + bar → the bar keeps its Stage 3 outcome, and the icon paints only with `icon:composition-with-bar`;
  - bar `showValue=1` + icon `showValue=0` → the descriptor pair is kept, and the text-suppression decision is marked unresolved (`icon:composition-bar-showvalue`);
  - icon + scale / DXF fill → the fill paints, and the icon only with its capability;
  - a matching p1 DXF `stopIfTrue` → no icon;
  - an icon with `stopIfTrue="1"` → `tainted` contains `ALL` until `icon:composition-icon-stopIfTrue`;
  - a `NoIcons` win over a lower matching icon → the lower icon is withheld until `icon:composition-noicons-slot`;
  - an unresolved p1 `stopIfTrue` stub → no icon.

  `visualSlotOf` returns `'iconSet'` for an icon rule and an icon stub. All Stage 2/3 compose tests (`tests/xlsx-cf-scale-compose.test.ts`, `tests/xlsx-cf-bar-compose.test.ts`, `tests/xlsx-cf-bot-fixes.test.ts`) stay green unchanged.
- [ ] Extend `wire measures match the kernel payload walker exactly` (`tests/xlsx-cf-worker.test.ts:165`) with `cell-icon` and `cell-icon-bar`. `measureCfCellBytes` must equal `measurePayloadBytes` for both. `CF_ICON_CELL_BYTES` must equal the full-minus-empty difference, and `estimateCfPrepareBytes` must use it, with no hand constant (source-text assertion: no new integer literal on the icon arm).
- [ ] Add `'iconSet'`, `'iconPaint'` and `'iconGlyphs'` (created by Tasks 28 and 29, which re-run this test) to the file list in `excludes every CF implementation from both ordinary static closures` (`tests/xlsx-cf-lazy.test.ts:14–17`) (Lesson 7).
- [ ] Extend `tests/xlsx-cf-active-units.test.ts`. The largest production call with icons present (1,000 text-path cells and 256 paint-path cells, four icon rules each) spends ≤ `CF_LIMITS.activeUnits / 5` units. Re-run `npx tsx scripts/cf-refs/active-units.ts` and record the numbers in `docs/validation/xlsx-cf.md` §11 (Lesson 9).
- [ ] Run `npx vitest run tests/xlsx-cf-icon-eval.test.ts tests/xlsx-cf-icon-compose.test.ts tests/xlsx-cf-worker.test.ts tests/xlsx-cf-lazy.test.ts tests/xlsx-cf-statistics.test.ts tests/xlsx-cf-active-units.test.ts tests/xlsx-cf-bar-compose.test.ts tests/xlsx-cf-scale-compose.test.ts` to record RED, implement, then run GREEN. Run tsc and the eager gates. Commit (`feat(xlsx-cf): evaluate and compose icon sets`), then report and review.

### Task 27: Icon wire, worker grant, pending pages and whole-sheet invalidation (Claude Sonnet; review opencode)

**Files:**
- Modify:
  - `src/xlsx/view/pagedSource.ts`: `hasIcon`/`perCellIcon` and the `icon` omit family in `prepareCfRemoteOnce`, parallel to `:2000–2005` and `:2174–2177`;
  - `src/xlsx/cf/text.ts`: `CopiedState.icon`, `wireIcon`/`descriptorOfIcon` beside `wireBar`/`descriptorOf` (`:117,124`), `assembleScope`, and `prepareCfWireScope`;
  - `src/worker/viewerRemoteCf.ts`: icon bytes in `RemoteCfCache` adoption accounting.
- Create `tests/xlsx-cf-icon-worker.test.ts`. Extend `tests/xlsx-cf-live.test.tsx`.

**Interfaces:**
- **Consumes:** `CfCellStateWire.icon`, `CF_ICON_CELL_BYTES`, `measureCfCellBytes` and `CF_ICON_SET_NAMES` (Tasks 25–26); `measureCfWalkedBytes`, `measureCfPayload`, `admitReply`, `RemoteCfCache`, and `cfRulesVersion`.
- **Produces:**
  - icon states on the wire under the receiver grant, omitted when they do not fit (S4-Q);
  - `CopiedState.icon`, so the paged main thread sees the same descriptor;
  - the icon-only whole-sheet invalidation test.

- [ ] Write `icon payload fits the grant or is omitted`. A 1,000-cell viewport's icon bytes are priced from `CF_ICON_CELL_BYTES` before the payload is built.
  - At grant − 1, base cells and bars return, every `icon` field is omitted, and `cf-grant-required` ("…icon sets omitted…") is present only if it fits.
  - At exactly the grant, the reply carries every icon.
  - `admitReply` never replaces the reply.
- [ ] Write `wire icons equal the resident descriptor`. For a custom set with `NoIcons` under injected evidence, each decoded wire `icon` (with `set` → `CF_ICON_SET_NAMES[set]`) deep-equals the resident `visual.iconSet`. A wire `set` outside 0..20 decodes to no icon, never to a guessed set.
- [ ] Write `the main thread consumes no per-rule icon data (S4-Q, owner-accepted)`.
  - After a worker read of a sheet with a custom set carrying `NoIcons`, the adopted `RemoteCfCache` entry and every `CopiedState` hold only `CfCellStateWire` fields: no `iconSet` definition, no thresholds, no `cfIcon` list, and no `CfSharedDefinition` for icons.
  - A source grep in the test asserts no `definition.custom`, `definition.thresholds` or `kind === 'iconSet'` read in `src/xlsx/cf/text.ts`, `src/worker/viewerRemoteCf.ts`, `src/xlsx/view/renderViewport.ts` or `src/components/`.
- [ ] Write `CF_ICON_SET_NAMES is an immutable wire table`. A snapshot pins its 21 entries in order, so a reorder fails the test and is a breaking wire change. A wire `set` of −1, 21 or 2.5 decodes to no icon.
- [ ] Write `the advertised requiredBytes of a refused icon read carries every icon` (retry at exactly that grant) and `clone copies of icon states are counted on both sides until adoption`.
- [ ] Write `max edit re-buckets every icon of the scope in one generation`. With the resident and worker shells at g1 (max 100, v=60 in bucket 1 under `percent` 33/67), commit max 50 at g2 while preparation is pending: no cell paints its g1 icon at g2, and v=60 gets its g2 bucket.
- [ ] Write `a CF generation change with the lease kept invalidates every tile of an icon-only sheet (Lesson 5)`. It follows `tests/xlsx-cf-live.test.tsx:482` on a 60-row sheet whose only rule is an icon set. Every tile of the sheet invalidates, and no tile of another sheet does.
- [ ] Write `unimported scope pages withhold icons then repaint once`. The page shows no icon, no inset and no permanent diagnostic, then exactly one invalidation on arrival.
- [ ] Write `no-icon sheets do zero icon work`. A bar-only sheet's wire carries no `icon` key, and its byte estimate equals the Stage 3 estimate.
- [ ] Run `npx vitest run tests/xlsx-cf-icon-worker.test.ts tests/xlsx-cf-live.test.tsx tests/xlsx-cf-bar-worker.test.ts tests/xlsx-cf-scale-worker.test.ts tests/xlsx-cf-worker.test.ts tests/xlsx-cf-lifecycle.test.ts tests/worker-package-bootstrap.test.ts` to record RED, implement, then run GREEN. Run tsc and the eager gates. Commit (`feat(xlsx-cf): transfer icon sets under receiver grants`), then report and review.

### Task 28: Original vector glyph table and decoder (OD-9(b)) (Claude Opus; review opencode)

**Dependencies:** Task 28 depends on the Task 24 glyph-key sidecar `refs/stage4-icon-glyph-key.json` and on OD-9, which the owner decided on 2026-11-14 as option (b), original vectors. It cannot start until Task 24's glyph-key capture has succeeded (`icon:pixel-calibration` and `icon:glyph-key` verified). **This is the only task whose content depends on OD-9.** Every other task consumes only `iconGlyph`/`drawIconGlyph` and treats `undefined`/`false` as withheld.

**Files:**
- Create `src/xlsx/cf/iconGlyphs.ts`, holding original, hand-authored vector glyphs (one per glyph class, drawn to the measured box and dominant colours), and `tests/xlsx-cf-icon-glyphs.test.ts`.
- No crop, bitmap or traced Microsoft artwork enters the repository's source or package. The sidecar crops stay test fixtures, used only for comparison.
- Update `docs/validation/xlsx-cf.md` with the glyph provenance: original artwork, the author, the approximation tolerance, and the `pixel-approximated` label.

**Interfaces:**
- **Consumes:** `refs/stage4-icon-glyph-key.json` (crops, glyph classes, box px), `CfBarScale`, `paintMemoryFor`, and `createCfWorkBudget`.
- **Produces:** the Task 28 block of **Produces**: one table, one decoder and one per-batch, scale-independent admission of the glyph classes the batch uses, charged before allocation.

- [ ] Write `every (set, iconId) maps to the measured glyph class`. `CF_ICON_GLYPH_CLASSES` equals the sidecar's classification for all 75 pairs. Pairs that share a crop share one table entry (deduplicated), and the entry count equals the sidecar's class count.
- [ ] Write `unknown or unverified glyphs are withheld`. `iconGlyph('3Arrows', 3)` is `undefined`. A pair whose `icon:glyph-<set>-<id>` capability is unverified returns `undefined`, and so does `NoIcons`, which is drawn as nothing by the caller.
- [ ] Write `glyph decode is charged and bounded`.
  - Admitting every class the sidecar records (its class count, not a plan-authored number) stays within the cap, and `dispose` returns the ledger to baseline.
  - Drawing an admitted glyph at each of the six zoom/DPR pairs allocates nothing (ledger unchanged).
  - A refused charge returns `false`, draws nothing, and leaves the ledger at baseline.
  - Decoding spends work units, not wall time.
- [ ] Write `glyph table bytes are measured`. The encoded table's byte size, as read from the built `cf` chunk by `OV_MODULE_MAP=1 bun run build`, is recorded and asserted ≤ 16 KiB gzip (a vector table; the cap is stated here and reported in Task 30). Once Task 28 has measured the table, the assertion tightens to that measured size plus 10 % (Ruling S4-GLYPHCAP, plan r2 review); 16 KiB is only the ceiling before measurement. No glyph bytes appear in the eager closure (checked through the `tests/xlsx-cf-lazy.test.ts` file list).
- [ ] Write `original glyphs approximate the measured box and colours (OD-9(b))`.
  - At 100 %/DPR 1, each glyph's drawn ink bounding box equals the sidecar's measured box within ±1 px.
  - Its dominant colour per region equals the crop's within ±8 per channel.
  - The result is labelled `pixel-approximated`, never `pixel-measured`.
  - A glyph failing either check makes `iconGlyph` return `undefined` for its pairs (withheld).
- [ ] Write `the glyph table ships no Microsoft artwork`. A source-text assertion over `src/xlsx/cf/iconGlyphs.ts` finds no `data:` URI, no base64 run longer than 64 characters, no PNG/`ImageData` construction, and no import from `tests/fixtures/`. `CfIconGlyph.source` is `'original'` for every entry.
- [ ] Run `npx vitest run tests/xlsx-cf-icon-glyphs.test.ts tests/xlsx-cf-lazy.test.ts`, tsc and the eager gates. Commit (`feat(xlsx-cf): add the lazy icon glyph table`), then report and review.

### Task 29: All-view icon painting, text inset, `showValue`, accessibility and copy/find (Claude Opus; review opencode)

**Files:**
- Create `src/xlsx/cf/iconPaint.ts`, `tests/xlsx-cf-icon-render.test.ts` and `tests/xlsx-cf-icon-a11y.test.tsx`.
- Modify:
  - `src/xlsx/cf/render.ts`: `CfPaintBatch.iconAt`/`paintIcons`/`textInsetAt`, and `hidesTextAt` over both visuals;
  - `src/xlsx/render.ts`: beside `:913` (inset into `layoutCellBody`) and `:928` (`paintIcons` after `paintBars`), plus the new `inset` parameter of `layoutCellBody`;
  - `src/xlsx/view/renderViewport.ts`: `createCellLayoutHost` (`:29`; its `layoutBody` at `:34`), placements carrying `textInset` (`:245–246`), and `paintIcons` beside `:305`;
  - `src/xlsx/view/textLayout.ts`: `CellLayoutHost.layoutBody` `inset`, `PlacedCell.textInset`, and `layoutKey`;
  - `src/xlsx/view/formatWorkspace.ts`: the optional `EffectiveFormatScope` methods;
  - `src/xlsx/cf/text.ts`;
  - `src/xlsx/cf/describe.ts`;
  - `src/components/xlsx/InteractiveSheetDoc.tsx`: forward the new scope methods beside `:1434–1436`.
- Extend `tests/xlsx-cf-print.test.ts`, `tests/xlsx-cf-export.test.ts`, `tests/xlsx-cf-copy-find.test.ts`, `tests/xlsx-cf-paged-a11y.test.ts`, `tests/xlsx-cf-oracle-conformance.test.ts` and `tests/browser/xlsx-cf.spec.ts`.
- Eager modules reach `iconPaint.ts`/`iconGlyphs.ts` only through `CfPaintBatch`/`EffectiveFormatScope` methods, never by import.

**Interfaces:**
- **Consumes:** `CfIconDescriptor`, Task 24 sidecars (`observedIcons`), Task 28 `iconGlyph`/`drawIconGlyph`, `paintCellLayout` phases, `layoutCellBody`, `residentBarHosts`/`frameBarHosts`, `barHidesText`/`barWithheld`, `describeCfWin`, `descriptorKey`, `getPaintables`, `renderPrintPage` (`render.ts:524`), and the sparkline RTL/snap convention (`sparklines/geometry.ts:95–97`, `PT_PX` at `:40`).
- **Produces:**
  - the Task 29 block of **Produces**;
  - one shared icon painter for static, interactive, paged, print and export;
  - one inset/hide predicate;
  - the S4-N descriptor;
  - `CfBarHost.textWidth` and `CfFrameHostSource.cell().textWidth`, threaded through `residentBarHosts`/`frameBarHosts`. Resident hosts take `textWidth` from the existing static `textMeasurer` (`render.ts:1225`) while the host is built. That measurement counts as layout work under the existing `admit()` byte/work budget, never as a fresh uncharged pass. An unknown width withholds the icon (plan r2 review).

- [ ] Write `icon geometry conformance (Lesson 1)`.
  - For every `pixel-measured` cell of every verified Stage 4 case, `layoutIcon` at the capture scale (1 device px per pt) on the sidecar's cell rect reproduces the box `leftPx`/`topPx`/`rightPx`/`bottomPx` within ±1 px, and the text ink shift exactly.
  - Every pending, refused or non-discriminating cell asserts `iconAt(...) === undefined` and `textInsetAt(...) === undefined`.
  - No expected px is computed by the function under test.
- [ ] Write `icon bucket conformance (Lesson 1)`. For every `pixel-classified` cell, production's `visual.iconSet` maps, through `CF_ICON_GLYPH_CLASSES`, to the sidecar's glyph class.
- [ ] Write `CF_ICON_RENDER_MODEL equals the measured sidecars`. The box variant, offsets, gap and clip rules equal `refs/stage4-icon-geometry.json` and `refs/stage4-icon-inset.json`. Without the capability, the matching branch of `layoutIcon` returns `undefined`.
- [ ] Write `icons paint in the fill phase in all five views`.
  - A recording context shows fill → bar → icon → content (the bar/icon order only once `icon:composition-with-bar` has settled z-order; until then no icon paints on a bar cell).
  - The views are `renderSheet`, `renderSheetViewport` (resident and forced-paged), `renderPrintPage` at `PRINT_SCALE = 2`, and the PNG/PDF export route.
  - The icon rect is identical in source space across the five.
- [ ] Write `inset and icon come from one predicate on one host (S4-L, Lesson 11)`. Each of these lays the text out un-inset, keeps it visible, and paints no icon:
  - `iconGlyph` returning `undefined`;
  - a glyph class refused at batch preparation (`admitIconGlyphs` returning `undefined`);
  - a `NoIcons` bucket;
  - a narrow cell without `icon:geometry-narrow-clip`;
  - a coarse frame on a non-active cell.

  In a coarse frame, the retained active cell alone is inset and painted. A mutant that insets from `iconAt` alone (ignoring the host) fails.
- [ ] Write `inset overflow decides identically on every leg`. While `icon:inset-overflow` is unverified, a left-aligned number whose inset text would not fit is laid out un-inset with no icon. The resident, worker and forced-paged legs all announce no icon for it, because each takes `textWidth` from its own layout or measurer. A host with an unknown `textWidth` withholds.
- [ ] Write `inset enters the layout key and the ### width`.
  - `layoutKey` of two otherwise equal placements differs when `textInset` differs.
  - A left-aligned number "12345" in a 40 px cell with a 20 px inset is laid out from x = 20 (the hypothesis run).
  - A number that fits without the inset but not with it follows `icon:inset-overflow`, or withholds the icon.
  - Toggling a rule from no-icon to icon within one session yields a layout-cache miss, never a stale hit.
- [ ] Write `showValue=0 hides canvas text only where an icon paints (S4-L)`.
  - With `icon:showvalue-canvas-suppression`, the content phase draws no glyph run, and the icon paints.
  - Without it, the text draws.
  - A withheld icon never suppresses text.
  - Bar plus icon with conflicting `showValue` keeps the text until `icon:composition-bar-showvalue` is verified.
- [ ] Write `tile and full render carry identical icon primitives at six scales`. At 50/100/175 % zoom × DPR 1/2, an icon cell straddling a tile edge has identical device-snapped rects in the full and tiled renders. A mutant that drops icons on non-first tiles fails (mirroring `tests/xlsx-cf-scale-render.test.ts:264`).
- [ ] Write `CF icons never expand static, print or export extents`. A whole-column icon set on `A:A` leaves the print canvas and export dimensions at the data extent.
- [ ] Write `icon cell announces a bounded positional label, without a tab stop (S4-N, Lesson 10)`.
  - `describeCfWin` returns `{ key: 'cf:iconSet', text: '圖示 2/3' }` in zh-TW and `'Icon 2 of 3'` in en for bucket 1 of 3, whatever the glyph class or `reverse`.
  - Bar + icon gives `{ key: 'cf:dataBar+iconSet', text: '資料橫條 72%、圖示 2/3' }` (zh-TW).
  - `NoIcons` adds nothing.
  - The descriptor is ≤ 128 units, with no RGB or set name.
  - The value text stays in the accessible name, and focus order is unchanged.
  - An icon win changes `descriptorKey`, so the presentation cache misses.
- [ ] Extend `tests/xlsx-cf-paged-a11y.test.ts`. The forced-paged semantic scope announces an icon exactly where the frame painter paints it. It decides through `frameBarHosts` + `iconWithheld`. The test's `CfFrameHostSource` supplies `textWidth` from the frame's laid-out record, and a source without it withholds. The resident, worker and forced-paged labels are identical for the same cells (S4-N, Lesson 10).
- [ ] Write `showValue=0 value still copies and finds`. Resident and worker copy and find return the formatted value text for a `showValue="0"` icon cell in near and far rows.
- [ ] Extend `tests/browser/xlsx-cf.spec.ts` with `Stage 4 icon sets paint and announce on every leg, and legs agree`, `Stage 4 icon sets keep painting at the six zoom/DPR pairs`, `Stage 4 icons under frozen panes`, `Stage 4 live max edit re-buckets the icons`, and `Stage 4 print capture paints icons`.
- [ ] Run `npx vitest run tests/xlsx-cf-icon-render.test.ts tests/xlsx-cf-icon-a11y.test.tsx tests/xlsx-cf-paged-a11y.test.ts tests/xlsx-cf-print.test.ts tests/xlsx-cf-export.test.ts tests/xlsx-cf-copy-find.test.ts tests/xlsx-cf-oracle-conformance.test.ts tests/xlsx-cf-bar-render.test.ts tests/xlsx-cf-render.test.ts` to record RED, implement, then run GREEN. Then run `npx playwright test --config playwright.viewer.config.ts tests/browser/xlsx-cf.spec.ts --project chromium`, and `bun scripts/cf-compare.ts --stage 4 --mode static` and `--mode print`, recording pass/refused/pending separately. Commit (`feat(xlsx-cf): paint icon sets in every view`), then report and review.

### Task 30: Stage 4 validation, resource audit and PR evidence (Claude Sonnet; review opencode)

**Files:**
- Create `tests/xlsx-cf-stage4.test.ts`.
- Extend `tests/xlsx-cf-oracle-conformance.test.ts` (Stage 4 counts) and `tests/browser/xlsx-cf.spec.ts`.
- Update `docs/validation/xlsx-cf.md`, the manifest statuses, `scripts/eager-baseline.json` (any recorded growth), `src/xlsx/cf/verifiedCapabilities.ts` (via `--write --publish-stage4` only) and `CHANGELOG.md`.
- Fix product regressions in separate focused commits, each with RED evidence and an opencode re-review.

**Interfaces:**
- **Consumes:** everything above, `auditStage4Manifest`, and the Stage 1–3 audits.
- **Produces:**
  - a Stage 4 audit row per case and variant: status, evidence kind, measured vs produced px and glyph class, mode, stamp, byte peak, and work units;
  - the generated `VERIFIED_STAGE4_CAPABILITIES`;
  - the PR 4 evidence.

- [ ] Write `every enabled icon capability has verified evidence` and `every unresolved icon shape refuses atomically on both paths`. The tests cover only the `icon:*` rows: manifest-format mutations stay in Task 24, and nothing is duplicated here. Run RED, then GREEN.
- [ ] Write `corpus iconSet files parse identically on both paths`, with `describe.skipIf` when `corpus/` is absent (as in `tests/xlsx-pivots-axes.test.ts:543`). For every inventoried file, the resident and streamed `CfSheetRules` deep-equal and nothing throws. The file count must equal the one recorded in `docs/validation/xlsx-cf.md` §4.
- [ ] **Dual gate (Lessons 1, 4).**
  - Run `bun scripts/cf-refs/audit.ts --stage 4` until it reports `stageComplete=true`, with every `icon:*` disposition in {`verified`, `unverified-nondiscriminating`, `refused-unsupported`, `supported-structural`} and none `unaccounted`.
  - AND `npx vitest run tests/xlsx-cf-oracle-conformance.test.ts` must be GREEN with no injection.
  - Then run `--write --publish-stage4` twice; `git diff --exit-code src/xlsx/cf/verifiedCapabilities.ts` after the second run proves idempotency.
- [ ] Run `npx tsc --noEmit -p .`, `npx tsc --noEmit -p tsconfig.scripts.json`, `npx vitest run`, `bun run gates:core` (which includes `OV_MODULE_MAP=1 bun run build`), `bun scripts/check-eager.ts`, `bun scripts/eager-budget.ts` and `git diff --check`. Record:
  - the exact summaries;
  - the eager gzip (strictly below 38,720 B);
  - the `scripts/eager-baseline.json` delta;
  - the `cf` chunk delta against PR 3, with the glyph table's share.
- [ ] Browser: run `npx playwright test --config playwright.viewer.config.ts tests/browser/xlsx-cf.spec.ts --project chromium`, plus `--project webkit` if it is installed.
  - Cover saved, interactive, resident and forced-paged views, the six zoom/DPR pairs, frozen panes, a live max edit, print capture, the icon descriptor, and the inset.
  - An unavailable engine is a named blocked gate.
- [ ] Run `bun scripts/cf-compare.ts --stage 4 --mode static|interactive|paged|print|export`, bound to the workbook and oracle hashes. Unverified fidelity (zoom ≠ 100 %, and OD-9(b) artwork) is labelled `*unverified*`. `pixel-classified`/`pixel-approximated` is never relabelled as measured fidelity.
- [ ] Run `bun scripts/cf-refs/audit.ts --stage 4` over:
  - a 1M-row synthetic `A:A` scope with 100 icon sets, bars and scales sharing one scan;
  - a refused percentile icon set with a `num` icon set still rendering;
  - whole-column extents isolation;
  - glyph-admission pressure with every class the sidecar records in one batch;
  - distinct-range pressure up to the 1,000,000-unit cap;
  - 20 open/read/cancel/close cycles.

  Pass when there are zero unowned live bytes at every transition, and zero icon work, glyph decodes or reads on a no-CF baseline.
- [ ] In `docs/validation/xlsx-cf.md`, publish one row per Stage 4 `[measure]` item: ruling, capability, evidence kind, status, and the exact refusal still in force. Confirm that the Stage 1–3 suites and capability sets are byte-unchanged (`git diff main -- src/xlsx/cf/verifiedCapabilities.ts` touches only `VERIFIED_STAGE4_CAPABILITIES` and the combined set).
- [ ] Update the changelog. The opencode review is recorded at `.superpowers/sdd/<plan>/task-30-review.md` with verdict APPROVE. Budget one extra PR-bot round for cross-path semantics (the Stage 3 lesson). Prepare the PR 4 evidence and leave the merge to the owner. Stage 5 refinement starts only after PR 4 merges.

### Lesson-to-task mapping

| Lesson (brief numbering) | Task | Test |
|---|---|---|
| 1 Oracle first + production conformance | 24, 29, 30 | `stage 4 icon cells are withheld until their evidence lands`; `icon geometry conformance (Lesson 1)`; `icon bucket conformance (Lesson 1)`; the dual gate |
| 2 Exact rendering or withheld | 25–29 | `no threshold evidence ⇒ unresolved`; `unknown or unverified glyphs are withheld`; `every unresolved icon shape refuses atomically on both paths`; withheld-only FINDINGS guard |
| 3 Non-discriminating ⇒ gate withholds | 24 | `non-discriminating icon variant keeps its gate closed (Lesson 3)` (`gte` off-boundary, `reverse` middle bucket, 2 px box hypotheses) |
| 4 `--write` idempotent + dual gate | 24, 30 | `audit --stage 4 --write is idempotent (Lesson 4)`; the Task 30 `--publish-stage4` dual-gate step |
| 5 Whole-sheet invalidation | 27 | `a CF generation change with the lease kept invalidates every tile of an icon-only sheet (Lesson 5)` |
| 6 Measurer-derived bytes, ±1e308 safety | 26, 27, 28 | `wire measures match the kernel payload walker exactly` (cell-icon rows); `icon evaluation stays finite at ±1e308 (Lesson 6)`; `icon payload fits the grant or is omitted`; `glyph table bytes are measured` |
| 7 Eager budget; CF out of eager/paged-worker | 26, 28, 30 | `excludes every CF implementation from both ordinary static closures` (+`iconSet`, `iconPaint`, `iconGlyphs`); `bun run gates:core` with `eager-baseline.json` |
| 8 Linux CI: skipIf + TZ pin | 24 | osacompile test under `it.skipIf(!hasOsacompile)`; conformance TZ pin |
| 9 Deterministic work units; fresh fallback budget | 26, 28 | `icon percentile copy reserves before allocation, and the fallback gets a fresh budget (Lesson 9)`; the active-units extension; `glyph decode is charged and bounded` |
| 10 A11y from semantics; three legs agree; paged scope uses the painter's host rule | 29 | `icon cell announces a bounded positional label, without a tab stop (S4-N, Lesson 10)`; the `xlsx-cf-paged-a11y` extension; browser `… legs agree` |
| 11 Coarse plans: text changes only on a painted host | 29 | `inset and icon come from one predicate on one host (S4-L, Lesson 11)` |
| 12 Excel canonicalises; Excel-authored x14 source | 24, 25 | the `stage4-icon-authored` orchestrator step; the marker tests diffed against resaved XML; `standalone x14 custom set with NoIcons resolves` |

### Stage 4 owner decisions

1. **OD-7 (as amended) and OD-8, binding.**
   - Icon evidence is pixel-only (S4-A). The object-model probe lifts nothing.
   - At other zoom/DPR values, the measured point geometry is scaled and snapped, labelled `*unverified*`. Buckets stay exact (S4-K).
2. **OD-9 (decided 2026-11-14): glyph source = (b), original vector icons.**
   - **The icons:** they match Excel's measured box, placement, inset and dominant colours, but not its artwork.
   - **Exact-or-withheld:** buckets, placement and inset stay exact-or-withheld. The artwork is `pixel-approximated` and labelled `*unverified*` (spec §8, "Unverified features ship"). Each glyph is withheld unless its box (±1 px) and dominant colours (±8) pass the Task 28 check.
   - **Provenance:** no Microsoft artwork, crop, bitmap or trace ships, in the source or in the npm package. Captures remain test fixtures only.
   - **Why:** the capture route is fixed at 1 px/pt, so no Excel reference exists at 175 % or at DPR 2 under any option. Vectors stay crisp at the six zoom/DPR pairs, and they avoid redistributing Microsoft artwork in the MIT package.
   - **Rejected:** (a) measured Excel bitmaps (a licensing issue, and soft at high DPR); (c) withholding glyphs entirely, or a hybrid.
   - **Scope of the decision:** it shapes Task 28 only.
3. **OD-10 (decided 2026-11-14): positional labels.**
   - The label is 「圖示 b/k」 / "Icon b of k", from the bucket semantics (S4-N).
   - Named labels (`XlIcon` names, zh-TW UI strings, a `(set, iconId)` → name table confirmed by the owner) are a follow-up after PR 4.
4. **Ruling S4-Q (accepted 2026-11-14): no per-sheet icon definition transfer.** Per-cell ids ride the existing `CfCellStateWire`/`RemoteCfCache` lifecycle. Task 27's acceptance test and grep prove that the main thread consumes no per-rule icon data (spec §6.8 line 292).

### Self-review

- **Coverage.**
  - §3 line 84 defaults → S4-B/F, Task 25.
  - §5.10 thresholds, `gte`, `reverse`, custom and `NoIcons` → S4-B/C/D/E, Tasks 25–26.
  - Participation → S4-G.
  - §5.7 exclusivity and `stopIfTrue` → S4-I/J, Task 26.
  - §6.3 scope → S4-H.
  - §5.10 line 204 and §6.4 lines 261–266 (inset, `showValue`, fill phase, tile seams, zoom) → S4-K/L, Task 29.
  - RTL → S4-M.
  - §6.7 → S4-N, Task 29.
  - §6.8 including line 292 → S4-Q, Task 27.
  - §3/§4 dialect and refusals → S4-B/C/O/P, Task 25.
  - §7 and §11 line 350 → Global Constraints delta, S4-R, Tasks 28, 30.
  - §8 → S4-A, Task 24.
  - §9 items 4/6/10 → Review Focus 1–5, 8, 10.
- **Interfaces.** Every Consumes and architecture citation was checked by printing the cited line on `bdc2b217`, and eight wrong line numbers were corrected. The check was a hand-built list, not an automatic extraction of every citation in the prose. Round 1 fixes re-checked every changed citation: `renderViewport.ts:29/:34`, `cf-compare.ts:197`, conformance `:645–646`, `cf/render.ts:180–182`, `dataBarPaint.ts:139`, `viewerRemoteCf.ts:147` and `contracts.ts:409`. Every Produces name was grepped absent and is marked with its task. `visual.iconSet` has no consumer, so replacing it in place is safe.
- **Main-branch facts that change the plan.**
  - `NoIcons` is refused today (`parse.ts:151–155,1490–1492`).
  - The icon object model is unreachable by construction (sdef).
  - Layout caching keys on `cfKey`, but the inset must also enter `layoutKey`.
  - `resolveScaleThresholds` cannot take 4–5 points.
- **Proportion.** Seven tasks against Stage 3's six. Three things are new compared with Stage 3: the glyph key and classification (Task 24), the text inset through layout (Task 29), and the OD-9(b) glyph table (Task 28).

## Stage 5: general formulas and cross-kind completion (PR 5, one PR per OD-11, decided)

Refined from spec draft 4 (§3 x14 bullet at lines 76–79; §4 row 5 at line 113 and the fail-closed staging rule at line 116; §5.1 lines 125–126; §5.2 line 136; §5.3 lines 143–147; §5.7 lines 171–181; §5.8 line 186; §5.9 line 193; §5.11 lines 207–213; §6.3 lines 235–248; §6.5 lines 268–272; §6.6 lines 274–277; §6.7; §6.8; §7; §8; §9 items 5/8 at lines 329/332; §10 item 5 at line 342), the Stage 1–4 rulings, and the merged code on `main` at `fc61c6c6` (Stage 4, PR #79). The branch is `feature/xlsx-cf-s5`. Every `file:line` below was printed on `fc61c6c6`.

**Goal.** A formula-driven rule formats a cell exactly where Excel formats it, from the same values Excel's own calculation reads, or it formats nothing. The scope covers:
- `expression` rules (base `<formula>` and standalone x14 `<xm:f>`);
- non-literal `cellIs` operands (references, expressions, function calls);
- non-canonical `<formula>`/`<xm:f>` on text, blank and error rules, which Excel already evaluates as formulas (Stage 1 measured this: `cases.ts:794`, `:981–995`), and a canonical/divergent classification of `timePeriod` formulas, which main does not yet have (S5-G, Task 33);
- formula `cfvo` bounds and thresholds on colour scales, data bars and icon sets: `type="formula"`, and `num`/`percent`/`percentile` whose `@val` or `<xm:f>` is a reference or expression;
- standalone x14 colour scales (OD-3, deferred from Stage 2);
- x14 `activePresent="1"` activation predicates, conditionally (S5-J);
- references to other sheets of the same workbook;
- full cross-kind composition: every Stage 1–4 family under and over formula rules, including the unresolved-stub cases that Stages 1–4 could only withhold.

Formula outcomes paint in every view (page, interactive, paged/worker, print, PNG/PDF export) on the resident and paged paths, through the existing DXF, scale, bar and icon seams. Stage 5 adds no painter. None of the Stage 1–4 correctness is lost.

Out of scope, as before (spec §2): defined names, external workbooks, structured (table) references, and pivot `conditionalFormats`. Each is parsed, diagnosed and refused (S5-B).

**Architecture delta.** Stage 5 adds one evaluator input, a CF formula adapter, and one data dependency, the formula footprint. It adds no new rule family, no new visual slot, no new wire field per cell, no new lazy chunk and no painter. Formula outcomes reach paint as the booleans, literals and numbers the Stage 1–4 evaluators already consume.
- **Parser stubs to replace (`src/xlsx/cf/parse.ts`).**
  - `expression` ends in `finishUnresolved('cf-unsupported-rule', withDxf())` (`:1526–1527`).
  - Non-literal `cellIs` operands fail `parseLiteral` (`:274`) and finish unresolved (`:1449–1451`). The text operators on `cellIs` finish unresolved too (`:1439–1440`).
  - Non-canonical text and blank/error formulas finish unresolved (`:1470`, `:1482`). They are detected by `canonicalTextFormula`/`canonicalBlankFormula` (`:410`, `:431`), whose anchor test is `refsMatchAnchor` (`:341`). **Main-branch fact (round-1 correction):** `parse.ts:1489` refuses a `timePeriod` whose `<formula>` sits in a foreign XML **namespace**; it does not classify the formula text. A same-namespace divergent formula (`<formula>A2=TODAY()-3</formula>` with `timePeriod="today"`) resolves today as `{ kind: 'timePeriod', period: 'today' }` through the attribute (Ruling R1c, comment at `:1490`; pinned by `tests/xlsx-cf-parse.test.ts:257–262` and `tests/xlsx-cf-paint-fixes.test.ts:165`). R1c has no oracle row, so this is a fail-open against spec §5.3 line 143, not a stub; Task 33 fixes it. `duplicateValues`/`uniqueValues` with any formula are unresolved at `:1511–1512`.
  - Colour scales: every x14 scale is refused (OD-3, `:767–769`), and so are `activePresent` (`:772–774`), `type="formula"` (`:807–809`) and a reference or expression `@val` (`:834–842`).
  - Bars: `barBound` refuses `type="formula"` (`:982–983`), and the activation predicate is refused at `:1030–1032`.
  - Icons: `iconThreshold` refuses `type="formula"` (`:1225–1226`), and the activation predicate is refused at `:1250–1251`.
  - Shared: `cfvoLiteral` (`:960–969`) classifies a reference or expression as `cf-unsupported-rule`. `activePresent` on a non-visual rule is `cf-invalid-rule` at `:1413–1414`, and its non-boolean spelling at `:1374–1377`. The single-range anchor is `anchorOf(ranges[0])` (`:1419`), and multi-range anchoring is undefined.
  - Formula text is already bounded: `readFormulas` (`:1556`) caps children at `CF_LIMITS.formulaChildren` (3), text at `textUnits` (8,192) and tokens through `countFormulaTokens` (`:451`). The streamed capture keeps one overflow witness (`cfChildLimit`, `:2268`). The raw formula strings already ride on `CfRuleOrigin.formulas` (`model.ts:14`).
- **Formula engine (`src/xlsx/formula/`).**
  - `parseFormula` (`parser.ts:859`) over `tokenize` (`lexer.ts:149`) yields the `AstNode` union (`types.ts:206`), which includes `name`, `table`, `ref3d`, `union`, `intersect`, `spill`, `implicitIntersect` and `arrayConst`.
  - `evaluateNode` (`evaluator.ts:717`) and `evaluateFormulaInternal` (`:1100`) evaluate against `EvaluationContext` (`types.ts:92`). The context supplies `getCellValue`/`getRangeValues` (`:116`, `:125`), `semantics: ResolvedSemantics` (`:134`, whose fields at `:279` are `dateSystem`, `timeZone` and `epochNowMs`, not the spec's `zone`/`clock`), `currentCell`, `unsupportedFeatures` (`:105`) and `materializationBudget` (`:152`; type at `:90`; charged by `chargeMaterialization`, `evaluator.ts:179`).
  - **Main-branch fact (changes the adapter).** Conditional aggregates need `ctx.references`. Without it `COUNTIF`/`SUMIF` return `conditional-reference-unavailable` (`functions/cond.ts:159`, `:350`), and a whole-column area takes `sparse-population-unavailable` (`functions.ts:117`). Both are typical CF formulas (`COUNTIF($A:$A,A2)>1`). The adapter therefore builds `ReferenceServices` through `createReferenceServices` (`refs.ts:810`), with deps `ReferenceServiceDeps` (`refs.ts:50`) over a snapshot-backed `SparseCellStore` (`refs.ts:43`), as the workbook evaluator does at `workbook.ts:1693`. `definedNames: []` and `tables: undefined` make names and structured references unavailable, never known-empty.
  - **Main-branch fact (corrects the Stage 1 inspected-facts table, row `formula/types.ts…`).** `translateSharedFormula` is defined at `shared.ts:386` over `translateAst` (`shared.ts:270`). `workbook.ts:938` is only its call site. `translateAst` turns an out-of-sheet coordinate into `#REF!`, which is one of two competing hypotheses for CF (S5-C). It also allocates a fresh tree per call.
  - `FUNCTIONS` (`functions.ts:508`) is mutable through `registerFormulaFunction` (`:969`). An unknown function is recorded through `markUnsupported` (`evaluator.ts:1004`, helper at `:499`), but only when the branch actually runs, so lazy `IF` branches record nothing. **Not registered on main:** `SEARCH`, `FIND`, `EXACT`, `ISEVEN`, `ISODD`, `INDIRECT`, `OFFSET`, `SUMPRODUCT`, `RAND` and `CELL`. Excel's own synthesised text formulas (`NOT(ISERROR(SEARCH(…)))`) are therefore evaluable only through the Stage 1 canonical-pattern path.
  - **Main-branch fact (round-1):** the public boundary `evaluateFormulaInternal` turns a top-level referenced blank into `0` (`evaluator.ts:1132`), and a bare range at the root is projected to a scalar inside `evaluateNode` (`scalarOfReference`, `refs.ts:672`). Neither may sit in front of the CF truth gate (S5-D).
  - **Main-branch fact (round-1):** sparse aggregate traversal does not charge `MaterializationBudget`. `ResolvedRefImpl.populated` (`refs.ts:200`) scans the whole `store.stored()` once per area occurrence and then sorts natively. `SUM(A:A)+SUM(A:A)` over 50 cells performs 100 reads with `used=0` (reviewer counterexample). An admitted function can also record an unsupported argument profile and still return a value through `IFERROR`: `COUNTIF(A:A,TRUE)` records `conditional-boolean-criterion-unverified` (`functions/cond.ts:175`). Unsupported features, issues and cycles surface through `unsupportedFeatures`, `reportFormulaIssue` and `hasCycle` (`types.ts:105,162,166`).
  - `TODAY`/`NOW` read `semantics.epochNowMs`/`timeZone` (`functions/datetime.ts:321–328`, registry at `:616–617`). `recomputeDependents` (`workbook.ts:627`) throws under the cached policy (`:632`). `decideCellCalculation` is at `calc-policy.ts:72`, and `workbookValuePolicy` at `value-policy.ts:11`.
- **Evaluation (`src/xlsx/cf/service.ts`).**
  - `createCfService` (`:71`) evaluates each cell in `evaluate` (`:214–325`). Non-visual rules go to `evaluatePredicate` (`:296–297`, `predicates.ts:461`), whose `cellIs`/text/blank/error/`timePeriod` arms (`evaluateCellIs` `:80`, `evaluateText` `:145`, `evaluateBlanksGated` `:222`, `evaluateErrors` `:238`, `evaluateTimePeriod` `:342`) take a `CfLiteral` (`model.ts:10`).
  - Visual thresholds resolve per rule: `resolveScaleThresholds`, `resolveBarThresholds`, `resolveIconThresholds` (`:245`, `:273`, `:291`). The unit budget is `work()` (`:101`), and the per-cell page loop is at `:326–336`.
  - Stage 5 adds a per-batch **dependency snapshot**, built before the page loop, plus one formula arm per consumer.
  - The `statistical()` predicate (`:18`) and `estimateCfPrepareBytes` (`:37`) key on definition kinds. A new formula arm must be added to both deliberately, or it is silently priced at zero.
- **Value sources.**
  - Every `CfValueSource` (`model.ts:98`) is single-sheet: `residentSource` (`render.ts:68`), `storeCfSource` (`render.ts:129`), `createResidentCfSource` (`statistics.ts:1143`) and the paged source built in `readConditionalFormatting` (`pagedSource.ts:897`, stamp at `:982–990`).
  - Cross-sheet reads need a new workbook-scoped seam. The paged worker already reads other sheets' pages for charts: `fillCharts` (`pagedSource.ts:1198`) groups cells by `workbook.store.pages(sheetId)` and adds an unimported sheet to `pending` (`:1231`).
- **Generations and invalidation.**
  - `CfStamp` (`model.ts:5`) carries `workbookValueGen`. Every construction site fills it with the **host sheet's** `valueGen`: `InteractiveSheetDoc.tsx:1168`, `storeCfSource` (`render.ts:131`), `pagedSource.ts:985`, and `viewerSession.ts:3329,3590,3604,3627`.
  - `sameCfStamp` (`cache.ts:5`) and the remote cache's comparisons (`viewerRemoteCf.ts:105,140,614,650`) therefore cannot see a referenced sheet change. The resident live key `cfLiveKey` (`InteractiveSheetDoc.tsx:1083`) keys on the host `valueGen` only.
  - A remote pivot commit invalidates CF only for the committed sheet (`invalidateRegionStyles(sheetId)` → `invalidateCf`, `viewerRemoteSource.ts:706,414,417`). `RemoteCfCache.load` checks only the host sheet's `loadState` (`viewerRemoteCf.ts:514`). Waiting entries retry on every `notify` (`wakeWaiting`, `viewerRemoteSource.ts:784`).
  - **Main-branch fact.** The plan's Stage 1 Task 10 promised an "external source change invalidates CF" coordinator for Stage 5. No such symbol or test exists on main: grep finds neither, and `tests/xlsx-cf-live.test.tsx` has no test of that name. Stage 5 builds the cross-sheet path itself (S5-K) and cites only what exists.
- **Clock.**
  - The interactive view advances `clockGen` at local midnight (Ruling M2, `InteractiveSheetDoc.tsx:566–576`).
  - The static, text and worker preparations each read `Date.now()` themselves: `render.ts:342`, `text.ts:424` and `pagedSource.ts:997`. `text.ts:435` and `viewerSession.ts:3329` hardcode `clockGen: 0`.
  - Stage 1 published the `timePeriod:*` clock as measured in the Asia/Taipei UTC/local window (`docs/validation/xlsx-cf.md:72`). Formula `TODAY()` must use that same clock (§5.11), threaded identically on every leg (S5-M).
- **Calculation settings.** `calcMode` lives on `XlsxDocument.calc` (`src/xlsx/types.ts:247`, built at `parse.ts` of the workbook, `src/xlsx/parse.ts:330–340`). The paged import reads only `fullCalcOnLoad` (`src/xlsx/storage/pagedImport.ts:196–198`), and `PagedWorkbook` (`pagedImport.ts:71`) carries no calculation mode. A sheet-only preparation therefore cannot know the calc mode or the document value policy (S5-L).
- **Overlay-aware stores.** Each paged sheet's canonical store is a `PagedXlsxSheetAdapter` (`pagedSource.ts:110`, live `stamp` at `:132`), registered per workbook in `workbookAdapters` (`:88`, `:125`) and created on demand by `storeFor` in `createPagedWorkbookSource` (`:2498`, `:2513`); the resident twin is `storeFor` at `src/xlsx/view/source.ts:999`. These stores carry pivot overlays and the matching generation; raw `workbook.store.pages` (the `fillCharts` pattern) does not.
- **Value policy.**
  - `prepareCfPaint` defaults `valuePolicy` to `'cached'` (`render.ts:322`), but `prepareCfTextScope` defaults it to `'engine'` (`text.ts:438`).
  - UI callers pass `doc.valuePolicy` (`src/render/paint.ts:425`, `sheetAssets.ts:259`). A direct API caller that omits it gets paint and text from different value authorities, which breaks Lesson 13.
  - Formula footprints read many more cells than a host, so S5-L binds one authority.
- **Blank synthesis.**
  - `canRuleMatchBlanks` (`src/xlsx/cfLoader.ts:100–105`, **eager**) returns true for `unresolved` + `expression` only. A resolved `expression` kind, a formula-operand `cellIs` or a resolved non-canonical predicate would silently stop blank-cell synthesis for `ISBLANK`-style rules.
  - Stage 5 changes this eager helper (Lesson 7 justification in the Global Constraints delta).
- **Oracle facts.**
  - `CfOracleCase.stage` already admits `5` (`scripts/cf-refs/cases.ts:253`). `audit.ts` refuses `--stage 5` (`:2206–2207`).
  - `CfPartMarker.formulas` (`cases.ts:95`) pins rule-level formulas. `CfPartMarker.scale` (`:97–100`) has no `f` and no dialect, and `bar.cfvo.f`/`icon.cfvo.f` exist.
  - The semantic route is the Stage 1 DisplayFormat readback: `buildColorProbeScript`/`buildDecorationProbeScript`/`buildTextProbeScript` (`capture.ts:213,192,175`) and `probeLiveExcelCellObservations` (`:587`). It yields `fillSource: 'excel-display-format'` (`cases.ts:144`). This is Excel's own evaluation of the rule against Excel's own calculation, so it is the primary oracle for every formula rule (S5-A).
  - Live expression stubs already exist in captured workbooks: `MOD(G2,2)=1` (`cases.ts:1949`, `priority:unresolved-expression-stub`, **published**), the scale stub (`:3478`, `scale:composition-unresolved-expression-stub`, **published**), `MOD(I2,2)=1` (`:4342`, `bar:composition-unresolved-stub`, refused) and `MOD(L2,2)=1` (`:5297`, `icon:composition-unresolved-stub`, refused; Excel matched `halted`, `docs/validation/xlsx-cf.md:1870`). `MOD` is registered.
  - The three `*-unresolved-shapes.rejected.xlsx` twins (Stages 2–4) carry **no `<cfRule>` at all**, inspected with `unzip -p`. Excel discarded the whole sheet's CF, so nothing is known about `activePresent` or formula cfvo in isolation. Stage 5 authors each such shape in its own workbook (Lesson 12).
- **Closures.** `src/xlsx/formula/*` is already in the index closure (`scripts/eager-baseline.json`, `indexModules`), and the engine is outside `EAGER_FORBIDDEN`/`PAGED_WORKER_FORBIDDEN` (`tests/helpers/eager-graph.ts:38,67`). The lazy `cf` chunk importing the evaluator adds no module to either closure. Task 32 proves this from the module map rather than assuming it.

**Global Constraints delta.** The Stage 1–4 Global Constraints, binding requirements and common gates apply unchanged. Stage 5 adds or re-binds:
- **Eager budget (Lesson 7).**
  - The eager bundle stays strictly below 38,720 B gzip (`scripts/eager-budget.ts:14`).
  - The CF formula adapter (`src/xlsx/cf/formula.ts`, `dependencies.ts`) stays in `EAGER_FORBIDDEN`/`PAGED_WORKER_FORBIDDEN`, and both join the file list of `tests/xlsx-cf-lazy.test.ts:14–17`.
  - The only eager or index-closure changes allowed (round-2 #3):
    - the widened `canRuleMatchBlanks` arms (`cfLoader.ts:100`);
    - `CfRuleOrigin.dependencySheets` read for the live key (`InteractiveSheetDoc.tsx:1083`);
    - `dependencyGen` at each stamp construction site;
    - the evaluation-context call sites in `src/render/paint.ts:425`, `src/components/xlsx/sheetAssets.ts:259`, `InteractiveSheetDoc.tsx` (text and paint scopes) and the worker metadata reply. Each site only passes `evaluation: cf.cfDocumentEvaluationOf(doc)` (or the worker's `PagedWorkbook` twin), where `cf` is the module the site already holds from the lazy loader (`ensureConditionalFormatting`, `cfLoader.ts:79`); the builder itself is lazy. No eager module imports a CF value.
  - **Byte estimate:** about 12–25 B gzip per call site, 80–130 B gzip in total, over the +16 B allowance. It is recorded as one `recordedChanges` entry in `scripts/eager-baseline.json` with this justification (the `recordedChanges` pattern already exists in that file), and the eager entry stays strictly below 38,720 B gzip. `bun run gates:core` must pass with the entry.
  - Stage 4 already spent about 1 KB of index-closure growth on parser metadata. Any Stage 5 growth beyond the allowance needs a written justification in the Task 39 report, or it moves behind the lazy loader.
  - The formula engine is already in the index closure. Task 32 asserts that `indexModuleSetHash` is unchanged by the adapter.
- **Exact or withheld (Lesson 2).**
  - A formula rule paints only when every capability its shape needs is verified: grammar, anchor, truthiness, function family, reference kind and calculation policy.
  - A formula that reads a value the snapshot does not hold (pending, out of footprint, unavailable under the cached policy, or past the budget) is **unresolved**. It is never evaluated against a substituted blank, `0` or `""` (spec §5.1 line 126, §6.3 line 248).
- **Evidence (Lesson 4).**
  - `verifiedCapabilities.ts` changes only through `bun scripts/cf-refs/audit.ts --stage 5 --write`, which must be idempotent.
  - The Stage 5 export is carried over verbatim unless `--publish-stage5` is passed. Only Task 39 passes it.
  - The gate is dual: the audit's `stageComplete=true` AND conformance is green under production evidence with no injection. A capability publishes only when the reading Excel matched equals production's reading, as in the Stage 4 publication gate (`CF_STAGE4_PRODUCTION_READINGS`, `docs/validation/xlsx-cf.md:1867`).
  - The Stage 1–4 exports stay byte-unchanged (S5-STUB).
- **Invalidation (Lesson 5).** Any CF generation change that keeps the lease invalidates the whole sheet, and that includes a change in a **referenced** sheet's value generation (S5-K). No old-generation formula outcome paints after the commit.
- **Bytes and numerics (Lesson 6).**
  - Dependency snapshot entries, compiled-formula nodes and per-cell translated trees are priced from the payload measurer (`measurePayloadBytes`, `src/xlsx/view/payloadBytes.ts:62`), never hand-counted.
  - Formula results feeding thresholds stay finite near ±1e308 through the existing `lerp`/`safeFraction` (`colorScale.ts:36,431`). A non-finite formula result is unresolved.
- **Bounded work (Lesson 9).**
  - `evaluateNode` cannot yield, and the engine's sparse traversal charges nothing. CF work is therefore charged at the CF-owned seams the engine must pass through, **before** the work happens (S5-N):
    - **per-call store (round-4 #1):** each call sees a restricted `SparseCellStore` view built from the snapshot index for that call's translated atoms only (`CfDependencySnapshot.viewFor`), never the shared snapshot, so a full-store enumeration visits only the call's own footprint. Building the view spends one unit per view cell;
    - every `read` spends one unit first; every element yielded by `stored()` spends one unit first, and `populated()` (`refs.ts:200`) enumerates the **whole view** once per traversal occurrence, so each occurrence costs `v` enumeration units, where v is the view's cardinality, separately from the matched-cell reads; entering `stored()` also spends `v·⌈log₂ v⌉` units for the native sort that follows;
    - every `getCellValue`/`getRangeValues` call spends one unit per cell before it reads;
    - the per-cell evaluation reserves its evaluator scratch bound (S5-N) before the call.
  - **Bounded synchronous calls, yields between them (round-2 #2):** each per-cell `evaluateNode` call is synchronous and runs under a **static per-call bound** computed before invocation: `nodes + v (view build) + Σ over traversal occurrences of (v enumeration visits + matched-cell reads + v·⌈log₂ v⌉ sort) + single-cell reads + text units + conditional-matcher transitions` (S5-B), with v taken from `CfPayloadMetrics.viewCells`, computed from the measured payload of the cells this call will read. A formula whose per-call bound exceeds `CF_LIMITS.checkpointUnits` (128, `limits.ts:49`) is refused at compile or at footprint time (`cf-resource-limit`), never evaluated. Because one call cannot exceed one checkpoint's worth of units, no call needs to yield inside the evaluator; the preparation loop's existing `work()` checkpoint (`service.ts:101–108`) runs between cells, awaits `hostTask()` on `'yield'`, and checks `signal.aborted` there. This satisfies spec §7's cooperative checkpoints ("yield every 128 steps under `AbortSignal`") at cell granularity. Store spends made inside a call that return `'yield'` still count (the budget contract at `limits.ts:274–275`); the yield is honoured at the next between-cell checkpoint, at most 128 units later.
  - **Why not a resumable evaluator:** suspending `evaluateNode` mid-traversal would need retained engine frames across an await and their ownership on cancellation; the engine's suspension machinery is a worksheet-recalculation contract, not a CF one. Ruling S5-HOIST moves absolute-range aggregates (`SUM($A:$A)`, `A2>AVERAGE($A:$A)`, `COUNTIF($A:$A,A2)>1`) out of the call into the chunked statistics scan, so the cap applies only to the per-cell remainder. What still stays plain: relative-range aggregates and non-hoistable criteria whose single-cell work exceeds 128 units. Task 39 reports how many oracle and corpus formulas the cap still refuses.
  - A refused spend throws the CF refusal token, which taints the evaluation (S5-E); the rule is unresolved (`cf-resource-limit`) on the rest of the batch, deterministically in row-major order. `MaterializationBudget` still bounds array materialization, but is never the only charge.
  - Wall time is never a refusal reason.
  - A fallback attempt, such as rebuilding a refused shared snapshot per rule, gets its own fresh per-attempt budget, and all attempts of one preparation draw from one preparation ceiling of 2 × `CF_LIMITS.activeUnits`, so fallbacks cannot multiply total work.
  - `scripts/cf-refs/active-units.ts` is re-run, so `activeUnits` stays at least 5× the largest production call with formula rules present.
- **Charge before allocating (Lesson 15).** Each compiled AST is reserved on the parse account before `parseFormula` runs (S5-B). The snapshot, its growth peaks (old plus new while it grows), each translated tree, evaluator scratch (populated-slot arrays, flattened argument arrays, memo maps; S5-N), the `ReferenceServices` store index and every decoded cross-sheet page are reserved on the batch's `CfMemory` before construction. Each has an admission-owned lifetime that ends at batch dispose.
- **Every view agrees (Lesson 13).**
  - Every preparation route receives one `CfEvaluationContext` (calculation mode, resolved value policy, clock), and its other-sheet values from one overlay-aware `CfWorkbookSource` (S5-K, S5-L, S5-P).
  - A view that lacks an input withholds identically. No path fails open (S5-P).
- **CI (Lesson 8).** osacompile and Excel tests use `it.skipIf(!hasOsacompile)` (`tests/xlsx-cf-oracle.test.ts:12`). Clock-dependent harnesses pin `process.env.TZ`, as the clock conformance test does (`tests/xlsx-cf-oracle-conformance.test.ts:1411`).

**New policy caps [call], not claims about Excel.** Task 34 implements and tests them at limit and limit+1. Task 34 calibrates them against the Task 31 fixtures and `active-units.ts`. A change after calibration needs a recorded review.

| Cap | Initial value / handling |
|---|---|
| Compiled formula | ≤ `CF_LIMITS.formulaTokens` (4,096) AST nodes. More refuses the rule (`cf-resource-limit`) at compile, before any translation. |
| Footprint per formula | ≤ 64 reference atoms after admission. Whole-column and whole-row atoms clamp to the referenced sheet's populated bounds, as statistics do (§6.3 line 243). |
| Dependency snapshot | ≤ 262,144 cells and ≤ 8 MiB (`CF_LIMITS.scratchBytes`) per preparation, holding **only the per-cell footprint** (the atoms the per-cell remainder reads after hoisting). Hoisted ranges are never copied into it; they are scanned by the hoist (S5-HOIST). Over the cap, the rules whose per-cell footprint does not fit are unresolved and the others still evaluate. |
| Traversal work | Every per-cell store read, `stored()` element and sort charge spends from the per-cell work budget before it happens. Which budget pays for what is fixed in S5-HOIST ("Budgets"). |
| Evaluator scratch | Reserved per call before invocation from `cfEvaluatorScratchBytes(f, payload)`, where `payload` (`CfPayloadMetrics`) is measured while the snapshot is built: per atom occurrence, the cell count, total string units and longest string. It prices nodes, populated-slot arrays, text intermediates and character arrays, and the conditional matcher's DP table (S5-B). Measurer-derived where the payload measurer can see the structure; the matcher's typed arrays at a conservative 8 B per DP cell [call]. A refused reservation means the call never runs. |
| Hoisted aggregates | One scalar or frequency map per `(rule, range, valueGen, dependencyGen)`, built by the chunked statistics scan under `scan: true` budgets and charged within `CF_LIMITS.statisticsBytes` (8 MiB) with its growth peak; over the cap the rule is `cf-resource-limit` (S5-HOIST). |
| Per-call work | ≤ `CF_LIMITS.checkpointUnits` (128) per synchronous evaluation; more refuses the formula (round-2 #2). |
| Materialization | One `MaterializationBudget` per preparation, `limit` = 1,000,000 cells, for array materialization only. |
| Referenced sheets | ≤ 16 distinct sheets per sheet's rules. More refuses those rules (`cf-resource-limit`). |
| Invariant results | A cell-invariant subtree (no relative reference and no `ROW()`/`COLUMN()`) evaluates once per preparation and is memoised under the batch's lease. A cfvo formula is memoised only under its measured scope (S5-H). |

### Stage 5 scope rulings

Each ruling names the spec clause it interprets, its fail-closed default and its cost if wrong. "Unresolved" means the spec §5.7 atomic fail-closed outcome:
- the rule's own contribution is withheld;
- an unresolved `stopIfTrue` rule withholds everything after it;
- any other unresolved rule taints its write mask for the rules after it (`compose.ts:239–245`).

A default never ships as a measured fact. Only the named capability lifts a ruling, through `hasEvidence` (`evidence.ts:88`). The capability prefix is `formula:*`. Formula bounds on visuals also use the family prefixes (`scale:*`, `bar:*`, `icon:*`).

- **Ruling S5-A: evidence authority (semantic first).**
  - **Spec tie:** §8 and §5.11, where every coercion and anchor is `[measure]`.
  - **Trusted source:** Excel's own DisplayFormat readback per cell (`buildColorProbeScript`/`buildDecorationProbeScript`, `capture.ts:213,192`), on the Excel-resaved workbook, after Excel has calculated it. The question for a formula rule is whether the format applied to this cell. DisplayFormat answers that from Excel's calculation, not from our evaluator. A pixel capture backs it up, using the Stage 1 `pixel-matched` fill comparison, and is required for print.
  - **Visual bounds:** a formula bound on a scale, bar or icon reuses the Stage 2–4 pixel evidence kinds unchanged. The formula only moves the threshold.
  - **Competing hypotheses (Lesson 3):** every case is designed so that each pair of hypotheses for a `[measure]` item predicts different outcomes at ≥ 2 designated cells per configuration (`replicationShortfall`, `verify.ts:5061`, Ruling S4-REPLICATE). A configuration with fewer agreeing decisive cells is `unverified-nondiscriminating`.
  - **No object-model reading of the formula itself:** Excel's resaved XML is diffed instead (Lesson 12). The formula text Excel writes back (canonicalised references, sheet quoting, sqref order) is what the marker pins.
  - **Cost if wrong:** a mis-designed case lifts a coercion. The defence is the two-cell rule plus a negative control per configuration.
- **Ruling S5-B: grammar admission and the function allow-list (OD-13, decided).**
  - **Spec tie:** §2 (names and external refs are refused), §5.11 line 210 (unresolved functions, external refs, defined names, cycles and budget refuse) and §6.2 line 233 (AST caps).
  - **Static, not lazy:** admission walks the parsed AST once at parse time (`compileCfFormula`), so it cannot depend on `unsupportedFeatures`, which records nothing for an untaken branch. It refuses the rule (`cf-unsupported-rule`) on any of:
    - a `name`, `table`, `spill`, `arrayConst`, `union` or `intersect` node;
    - a `ref3d` node (sheet runs);
    - a sheet qualifier containing `[` (external workbook);
    - a call whose upper-cased name is not in the frozen `CF_FORMULA_FUNCTIONS` table.
  - **Reservation before construction (round-1 #3):** `compileCfFormula` takes the parse-time `CfMemory`. Before calling `parseFormula` it reserves `countFormulaTokens(text) × CF_FORMULA_NODE_BYTES` (the token count already bounds the node count, `parse.ts:451`). A refused reservation returns `cf-resource-limit` and constructs nothing. The reservation is adopted into the rule's parse lease and released with it (parse failure, cancelled import, sheet disposal).
  - **Bound at compile:** the table also binds each admitted name to the handler captured at compile time. A later `registerFormulaFunction` cannot widen it.
  - **Allow-list (OD-13, decided 2026-11-16):** only names registered on main today (`functions.ts:508` and the `functions/*` tables):
    - `AND`, `OR`, `NOT`, `IF`, `IFERROR`, `IFNA`;
    - `ISBLANK`, `ISERROR`, `ISNA`, `ISNUMBER`, `ISTEXT`, `ISLOGICAL`;
    - `LEN`, `TRIM`, `LEFT`, `RIGHT`, `MID`, `UPPER`, `LOWER`, and `CONCAT` with scalar (single-cell or literal) arguments only;
    - `ABS`, `INT`, `MOD`, `ROUND`;
    - `ROW`, `COLUMN`;
    - `SUM`, `COUNT`, `COUNTA`, `AVERAGE`, `MIN`, `MAX`;
    - `COUNTIF`, `COUNTIFS`, `SUMIF`, `SUMIFS`, `AVERAGEIF`, `AVERAGEIFS`;
    - `DATE`, `YEAR`, `MONTH`, `DAY`, `WEEKDAY`;
    - `TODAY`, which is clock-gated (S5-M).
  - **Withheld by default:**
    - `NOW` (S5-M);
    - the lookup family (`INDEX`, `MATCH`, `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `XMATCH`);
    - dynamic arrays (`SEQUENCE`, `FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `TRANSPOSE`);
    - `TEXT`, `DATEVALUE` and `TIMEVALUE`, whose results depend on the locale;
    - every unregistered name (`SEARCH`, `FIND`, `EXACT`, `ISEVEN`, `ISODD`, `INDIRECT`, `OFFSET`, `SUMPRODUCT`, `RAND`, `CELL`).
  - **Payload-sensitive profiles (round-2 #1):** every admitted profile has a conservative bound computed before invocation, or it is refused. Text work is the risk: `textChars` (`evaluator.ts:508–511`) builds one array element per character, and `CONCAT`/`&` build intermediate strings. The snapshot records, while it is built and charged, the string units of every cell it holds; the per-call text bound is the atoms' total string units plus literal units, multiplied by the text-op node count plus one, priced for the simultaneous string and character-array lifetimes. A text bound over the per-call scratch cap (64 KiB) refuses the formula. `CONCAT` (or `&`) with any multi-cell argument is refused at compile, whatever its size, because its output size depends on the payload of a range (`TEXTJOIN` is not registered on main). `LEN(CONCAT(A1:A64))` is therefore refused at compile, before any read or allocation.
  - **Conditional-matcher profile (round-3 #2):** a non-hoisted `COUNTIF`/`COUNTIFS`/`SUMIF`/`SUMIFS`/`AVERAGEIF(S)` whose criterion is text runs the engine's wildcard matcher per range cell, a DP over `(pattern tokens + 1) × (text length + 1)` cells (`functions/wildcard.ts:137`), with typed-array allocations. Its profile bound is `Σ over the range's cells of (tokens(criterion) + 1) × (length(cell) + 1)` transitions, priced in both work units and scratch bytes, with the criterion's token count taken from the engine's own tokenizer (`tokenizeWildcardPattern`, `functions/wildcard.ts:41`) and lengths from the snapshot's `CfPayloadMetrics`. If the criterion is not yet known before the call (it is not a literal or a snapshot cell), the bound cannot be computed and the formula is refused. If the bound exceeds the 128-unit per-call cap, the formula is refused. A number criterion against numbers costs one unit per range cell.
  - **Reference-returning argument profiles (round-2 #4):** each `CF_FORMULA_FUNCTIONS` entry declares, per argument, `scalar` or `range`. A multi-cell reference (`range`, `wholeCol`, `wholeRow`) is admitted only in a `range` slot (the aggregate and `COUNTIF`-family range arguments). In any other position (a `scalar` slot, an operator operand such as `A2:A4>0`, or a value-returning branch of `IF`/`IFERROR`/`IFNA`) it is refused at compile until `formula:truth-range-nested` is measured, because the engine projects it to a scalar (`IF(TRUE,A2:A4,0)` and `IFERROR(A2:A4,0)` return a scalar on main with nothing recorded) before CF could inspect it.
  - **No engine changes:** Stage 5 adds no function to the shared engine. A function family paints only with its `formula:fn-<family>` capability (S5-E).
  - **Cost if wrong:** CF formulas using `SEARCH` or lookups stay plain. These include Excel's own divergent text formulas.
- **Ruling S5-SUMIF-SHAPE: effective read geometry (orchestrator ruling, round 5).**
  - **Spec tie:** §6.3 line 248 (callbacks read from a pre-admitted snapshot; a read outside it must never be served as data) and §5.7.
  - **The fail-open.** `SUMIF`/`AVERAGEIF` resize their sum range to the criteria range's rows × cols (`resizeTarget`, `functions/cond.ts:237–243`, applied at `:258–259`), so the cells they read are not the literal reference atom. A view built from literal atoms omits cells. With A1:A2 blank, B1 = 0 and B2 = 7, `SUMIF(A1:A2,"",B1)=0` is false on the full store but true on an atoms-only view. Codex's example.
  - **Admission:** a function whose reads extend beyond its literal reference atoms is admitted only when its effective geometry equals its literal geometry. For `SUMIF`/`AVERAGEIF`, a sum range (third argument) must have exactly the criteria range's rows × cols; any other shape is refused at compile (`cf-unsupported-rule`). With the sum range omitted, the engine reuses the criteria range (`cond.ts:253`), which is literal and admitted.
  - **Coverage rule:** every `CfFormulaFunction` declares, per argument, its **effective read geometry** (`literal`, or `resized-to: <argument index>`). `planCfHoist`, `cfFormulaFootprint`, the snapshot, `CfDependencySnapshot.viewFor`, `CfPayloadMetrics` and the hoist scans are all built from the effective geometry, never from the literal atoms alone. Whole-column and whole-row geometry is preserved as geometry (so the engine's absent-cell count, `view.ref.countAbsent()` at `cond.ts:270`, sees the true area) and bounded by population. `viewFor` validates, **before any sparse traversal**, that every effective area is fully covered by the snapshot's covered rects; an uncovered area refuses the call (`cf-formula-dependency`). Read-time taint stays as a second line of defence, but never the only one.
  - **Allow-list sweep (round 5), against `fc61c6c6`:**

  | Function | Implicit expansion? | Result |
  |---|---|---|
  | `SUMIF`, `AVERAGEIF` | Yes: sum range resized to the criteria range (`cond.ts:258–259`) | Admitted only with a same-shape sum range (or none); otherwise refused at compile |
  | `SUMIFS`, `AVERAGEIFS`, `COUNTIFS` | No: every range is read as written; mismatched shapes return `#VALUE!` (`cond.ts:265`) | Admitted with literal geometry |
  | `COUNTIF` | No: one range, no target | Admitted with literal geometry |
  | `SUM`, `COUNT`, `COUNTA`, `AVERAGE`, `MIN`, `MAX` | No | Admitted with literal geometry |
  | `ROW`, `COLUMN` | No value read (geometry only); a multi-cell argument is already refused in a scalar slot (S5-B) | Admitted |
  | `ISBLANK`, `ISERROR`, `ISNA`, `ISNUMBER`, `ISTEXT`, `ISLOGICAL`, text, math, date and logical functions | Single-cell or scalar arguments only | Admitted with literal geometry |
  | `INDEX` | Reference-returning; lookup family | Not admitted (S5-B) |
  | `SUMPRODUCT`, `OFFSET`, `INDIRECT` | — | Not registered on main; refused |
  | Dynamic arrays (`SEQUENCE`, `FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `TRANSPOSE`) | Output geometry depends on data | Not admitted (S5-B) |

  - **Cost if wrong:** `SUMIF` with a single-cell or differently shaped sum range stays plain.
- **Ruling S5-C: anchor and relative translation.**
  - **Spec tie:** §5.11 line 209.
  - **Single range:** the anchor is the top-left of the sole range, as main already assumes (`parse.ts:1419`). It needs `formula:anchor-single`.
  - **Multi-range `sqref`:** three hypotheses: (h1) the top-left of the first listed range; (h2) the top-left of the bounding box; (h3) each range's own top-left. The fixture orders `sqref` so the first range is not top-left (`D5:D8 B2:B4`). Until `formula:anchor-multi-range` selects one reading, a multi-range formula rule is unresolved. Excel's resaved `sqref` order is recorded first, because Excel may re-sort it on resave (Lesson 12). If it does, the reading is defined on the resaved order.
  - **Pinned axes:** a `$`-pinned axis never moves (`$A$1`, `$A1`, `A$1`), under `formula:anchor-mixed-abs`.
  - **Off-sheet shift:** a shift past the sheet edge has two hypotheses: (w1) wrap modulo 1,048,576 rows / 16,384 columns; (w2) `#REF!`, as `translateAst` does. Until `formula:relative-wrap` selects one, a translated reference that leaves the sheet makes the cell unresolved.
  - **New translator:** translation uses the new `translateCfFormula`, never `translateAst`. It takes the measured edge mode and charges its output tree before allocating (Lesson 15).
  - **Cost if wrong:** multi-range and edge-wrapping formula rules stay plain.
- **Ruling S5-D: result shape, projection and truthiness.**
  - **Spec tie:** §5.11 line 210.
  - **Shape before projection (round-1 #7):** the CF adapter never calls `evaluateFormulaInternal`, whose top-level blank→`0` (`evaluator.ts:1132`) would turn a blank result into a number. It classifies the root first:
    - a root `cell` reference reads the snapshot directly and keeps `blank` as `blank`;
    - a root `range`, `wholeCol` or `wholeRow` is a **range result**. CF projects it explicitly under the selected `range-*` reading (implicit intersection with the target row/column, top-left, or no-match), never through the engine's `scalarOfReference`;
    - any other root runs through `evaluateNode` in scalar mode with a CF-owned `nodeValues` memo; a `MatrixValue` result is a range result.
  - **Readings by observable outcome (round-1 #11):** hypotheses are defined by what the cell shows and whether lower rules continue, and indistinguishable readings are merged:
    - number: `number-nonzero-true` (0 ⇒ no-match);
    - boolean: `boolean`;
    - text: `text-no-match` (no format, lower rules continue; "false" and "error ⇒ no-match" are the same observable and are one reading) vs `text-true` (formats and, with `stopIfTrue`, halts);
    - blank: `blank-no-match` only. Treating a blank result as zero predicts the same observable under `number-nonzero-true`, so the two readings are merged (round-2 #7); `formula:truth-blank` still gates it;
    - error: `error-no-match` vs `error-halts` (an erroring `stopIfTrue` rule still halting lower rules);
    - range: `range-implicit`, `range-top-left`, `range-no-match`, discriminated by values that differ between the intersected cell and the top-left.
  - **Gating:** each kind needs its `formula:truth-<kind>` reading. Evidence for one kind never authorises another: number evidence alone leaves a blank or range result unresolved. `coerceToBoolean` (`evaluator.ts:382`) is never used as the fact.
  - **Cost if wrong:** rules returning text, blanks, errors or ranges stay plain.
- **Ruling S5-E: evaluation context and the values a formula reads.**
  - **Spec tie:** §5.11 lines 209–212, and §6.3 line 248.
  - **One value per cell:** the formula reads exactly the `CfLiteral` the host evaluation reads for that cell, the grid-displayed value (S5-L), converted once by `cfLiteralToElement`:
    - blank → `null`;
    - an error → its tagged `EvaluationError`;
    - a date serial → its number;
    - the rest as literals.
  - **Context:** `currentCell` = the target cell, `currentSheet` = the host name and `typedValues: true`. `semantics` comes from the sheet (`dateSystem`, `timeZone`) plus the S5-M clock. `references` is `createReferenceServices` over the snapshot store, with `definedNames: []` and `tables: undefined`.
  - **Taint, never coerce:** a read outside the snapshot, of a pending page, or of an unavailable value sets the evaluation's taint flag and returns `#N/A`. A tainted result is discarded as unresolved (`cf-pending-data` or `cf-formula-dependency`), whatever value it computed.
  - **Recorded issues discard the result (round-1 #1):** each evaluation gets a fresh `unsupportedFeatures` set, a `reportFormulaIssue` collector and `hasCycle = false`. If after the call any feature or issue was recorded, `hasCycle` is set, a materialization or work refusal happened, or the taint flag is set, the result is discarded as unresolved (`cf-formula-unsupported`, `cf-resource-limit` or `cf-formula-dependency`). This holds even when `IFERROR`/`IFNA` masked the error: `IFERROR(COUNTIF(A:A,TRUE),TRUE)` returns `true` on main while recording `conditional-boolean-criterion-unverified`, and CF treats it as unresolved, tainting its writes for lower rules.
  - **Self-reference:** a formula reading its own host cell is ordinary, since a CF rule writes no value. A referenced cell whose own worksheet formula is circular reads its grid value.
  - **Function families:** each needs `formula:fn-<family>` (logical, information, text, math, aggregate, conditional-aggregate, date, row-column). The evidence is the family's CF cells discriminating under DisplayFormat. Engine correctness itself stays covered by the engine's own tests.
  - **Cost if wrong:** a family whose CF behaviour is unmeasured stays plain.
- **Ruling S5-F: non-literal `cellIs` operands.**
  - **Spec tie:** §5.1 lines 125–126.
  - **Per-cell operand:** each operand formula is evaluated per cell under S5-C/S5-E, to a scalar converted back to a `CfLiteral` (`cfElementToLiteral`), then handed to the existing `evaluateCellIs` (`predicates.ts:80`). The Stage 1 operator semantics and their `cellIs:*` capabilities are therefore reused unchanged.
  - **Operand kinds:** `formula:cellis-operand-reference`, `-relative`, `-expression` and `-cross-sheet` gate them.
  - **Errors:** an operand that evaluates to an error refuses the rule on that cell. It never falls back to `0` or `""` (spec line 126).
  - **Text and blank operands:** these need `formula:cellis-operand-text` / `-blank`.
  - **Text operators:** the text operators on `cellIs` (`containsText`, `notContains`, `beginsWith`, `endsWith`; `parse.ts:1439–1440`) are **re-deferred**. Excel's dialog writes the standalone text types instead. `stage5-unresolved-shapes` records Excel's handling of a hand-authored `cellIs containsText`, and the shape stays `cf-unsupported-rule`.
  - **Cost if wrong:** `cellIs` with error operands stays plain.
- **Ruling S5-G: non-canonical text, blank, error and date formulas.**
  - **Spec tie:** §3 line 77 and §5.2 line 136.
  - **Measured fact:** Stage 1 measured that Excel evaluates the divergent `<formula>`, not `@type`/`@text`: `cases.ts:794` (`SEARCH("zz",J2)` overrides `@text="ab"`), and `:981–995` (`ISBLANK(D2)`, `NOT(ISBLANK(G2))`, `ISNA(J2)`).
  - **Default:** a non-canonical formula on these types compiles as an expression rule with the original `@type` kept for the write mask. It is evaluated by S5-B–E and needs `formula:predicate-divergent`.
  - **Canonical formulas:** these keep their Stage 1 predicate path (`canonicalTextFormula`/`canonicalBlankFormula`), so no measured Stage 1 outcome changes.
  - **Unregistered functions:** a divergent formula using an unregistered function (the `SEARCH` cells) stays unresolved under S5-B.
  - **`timePeriod` (round-1 #9):** main has no classifier (see the architecture fact). Task 33 adds `canonicalTimePeriodFormula(period, formulas, anchor)` (new), whose ten canonical patterns are read from the Excel-resaved Stage 1 clock workbooks (`stage1-timeperiod-today-clock`, `-2`), never written by hand. On both parser paths:
    - absent or canonical formula → the Stage 1 attribute predicate, unchanged;
    - divergent formula → unresolved (`cf-unsupported-rule`) until `formula:predicate-timeperiod-divergent` records which one Excel follows; if Excel follows the formula, it compiles as an expression (S5-B–E) with `sourceType: 'timePeriod'`. An unevidenced divergent formula never reaches the attribute predicate.
  - **R1c reconciliation:** the R1c tests at `tests/xlsx-cf-parse.test.ts:257–262` and `tests/xlsx-cf-paint-fixes.test.ts:165` are rewritten in Task 33 to expect unresolved for the divergent and off-by-one formulas. This is a deliberate Stage 1 behaviour change towards spec §5.3, recorded in the changelog; every Stage 1 clock conformance cell (Excel-written canonical formulas) must stay byte-unchanged.
  - **Duplicates:** a formula on `duplicateValues`/`uniqueValues` (`:1511`) is re-deferred. Excel writes none, and the shape stays unresolved.
  - **Cost if wrong:** divergent date rules stay plain.
- **Ruling S5-H: formula cfvo bounds and thresholds.**
  - **Spec tie:** §5.8 line 186 ("evaluated once per range"), §5.9 line 193 and §5.10.
  - **Default:** a cfvo of `type="formula"`, or `num`/`percent`/`percentile` with a reference or expression `@val`/`<xm:f>`, compiles once and evaluates to a number at the anchor (S5-C).
  - **Evaluation scope (round-1 #10):** on a single-range rule there is one scope, and the result is memoised per rule per preparation. On a multi-range `sqref` the two hypotheses are rule-wide (one value at the first range's anchor) and per-range (one value per range, at each range's own anchor). The `stage5-cfvo-scope` oracle discriminates them with ≥ 2 decisive cells per configuration. Until `formula:cfvo-scope-multi-range` selects one, **every multi-range rule with a formula bound is unresolved**. Once selected, memoisation is keyed by that scope (rule, or rule + range index).
  - **Feeding thresholds:** `num`/`formula` feed the threshold directly. `percent`/`percentile` feed the value p into the existing `lerp`/`percentileOrderStatistic` paths.
  - **Gates:**
    - `scale:thresholds-formula`, `bar:bounds-formula` and `icon:thresholds-formula` per family;
    - `formula:cfvo-reference` / `-expression`;
    - `formula:cfvo-relative`, because a relative cfvo formula's anchor is `[measure]`.
    - `scale:thresholds-num-val-reference` is already verified by Stage 2 (`docs/validation/xlsx-cf.md:252`), but production refuses it today (`parse.ts:834–842`). It now lifts, together with `formula:cfvo-reference`.
  - **Non-numeric results:** a cfvo result that is text, blank, boolean or an error is unresolved until `formula:cfvo-nonnumeric` selects a reading. The fixture's hypotheses are rule withheld, bound treated as 0, and bound skipped.
  - **Cost if wrong:** visuals bounded by cells stay plain.
- **Ruling S5-I: standalone x14 colour scales (OD-3 pickup).**
  - **Spec tie:** §3 line 75 and Ruling S2-K.
  - **Default:** a standalone `<x14:cfRule type="colorScale" priority>` with `<x14:colorScale>`/`<x14:cfvo><xm:f>` resolves through the Stage 2 point model with `scale:dialect-standalone`. The XML shape comes from `excel-authored-formulas.xlsx` (Lesson 12), which is authored with a bound on another sheet, the case S2-K expected.
  - **If Excel writes base:** if Excel writes such a scale in the base dialect instead, `scale:dialect-standalone` is recorded `refused-unsupported` with that reason, never retried, and x14 scales stay `cf-unsupported-rule`.
  - **Cost if wrong:** x14 scales stay plain.
- **Ruling S5-J: `activePresent="1"` activation (conditional pickup).**
  - **Spec tie:** §3 line 79 and §5.11 line 210.
  - **Authoring order:** Excel's dialog is tried first. If no dialog writes it, a hand-authored x14 visual with `activePresent="1"` and one rule-level `<xm:f>` is opened and resaved in its own workbook, one per visual kind.
  - **If Excel strips or repairs it,** or drops the rule: the outcome is recorded (Lesson 12, no retry), `formula:active-present-*` is `refused-unsupported` with that reason, and the shape stays `cf-unsupported-rule` permanently. That is the expected outcome, because the Stage 2–4 twins lost everything.
  - **If Excel retains it** and the capture discriminates (an activation false on ≥ 2 cells where the visual would otherwise paint): the activation is a per-cell S5-D truth gate ahead of the visual's own evaluation, under `formula:active-present-<kind>`.
  - **Unchanged refusals:** a missing or extra rule-level `<xm:f>` stays `cf-invalid-rule`, and so does `activePresent` on a non-visual (`parse.ts:1413`).
  - **Cost if wrong:** activated visuals stay plain, as today.
- **Ruling S5-K: cross-sheet values, pending pages and generations.**
  - **Spec tie:** §5.11 line 213, §6.3 lines 246–248, §6.5 line 272 and §6.6.
  - **Default:** a reference qualified by another sheet of the same workbook resolves through the new workbook-scoped seam `CfWorkbookSource`.
  - **One authoritative store (round-1 #6):** each dependency's values **and** its generation come from the same canonical, overlay-aware `SourceStore`:
    - paged worker: the sheet's `PagedXlsxSheetAdapter`, through `workbookAdapters` (`pagedSource.ts:88`) or, for a sheet never activated, created by the same `storeFor` path (`:2513`), so a first cross-sheet read builds the same store a later activation reuses; values through `storeCfSource` (`render.ts:129`) over that store, the generation from its `stamp` (`:132`);
    - resident interactive: `source.store(sheetId)` (`src/xlsx/view/source.ts:999`);
    - static, print and export: the document's sheets, which carry no overlay; a static leg mixing a host store with document sheets withholds cross-sheet rules.
    Raw `workbook.store.pages` is never used for CF values, because it does not carry committed pivot overlays.
  - **Recheck before publication:** after the snapshot is built and again before the preparation is adopted, each dependency store's stamp is re-read; any change cancels the preparation (`cf-stale-generation`), so a value read before a commit can never be published under the newer generation.
  - **Unknown or unimported sheets:** an unknown sheet name is `#REF!` under `formula:cross-sheet-missing`, else unresolved. An unimported page (`row > importedThroughRow` or status `'unknown'`) makes the cells reading it pending, with `pendingSheetIds` naming that sheet and no permanent diagnostic.
  - **Generations:** `CfStamp` gains `dependencyGen?: number` (absent ≡ 0): the sum of the referenced sheets' value generations, excluding the host. It strictly increases on any commit to any of them.
    - Every `source.stamp()` computes it live, so `CfPrepared.at` (`service.ts:134`) goes stale **synchronously** at a referenced sheet's commit, exactly as it does today for the host. `sameCfStamp` and the four `viewerRemoteCf.ts` comparisons include it.
    - The worker reply lists the referenced sheets in a new `CfReadResult.dependencies` field (each a `{ sheetId, valueGen }` pair). The main `RemoteCfCache` entry is current only while each listed identity keeps that `valueGen`.
    - `viewerRemoteSource` invalidates the CF of every sheet whose cached reply lists a committed sheet (beside `:706` and `:752`).
  - **Unknown generation:** a referenced sheet whose value generation the worker cannot name refuses those rules (`cf-dependency-generation-unknown`).
  - **Gate:** `formula:cross-sheet-reference`.
  - **Cost if wrong:** cross-sheet rules stay plain or pending. No stale paint, because the stamp check is synchronous.
- **Ruling S5-L: calculation and value policy (OD-12, decided).**
  - **Spec tie:** §5.11 line 212.
  - **Default (OD-12 decided 2026-11-16: option (a)):**
    - Formulas read the grid-displayed value of every cell: the cached value or the engine's value, whichever the grid shows under `doc.valuePolicy`.
    - Never `recomputeDependents`.
    - Under `'cached'`, a footprint cell that `isValueUnavailable` (`view/values.ts:50`) taints the evaluation (S5-E).
    - A CF formula still evaluates under `'cached'`, because `<cfRule>` stores no result to show.
  - **Manual calculation:** under `<calcPr calcMode="manual"/>`, formula rules are unresolved until `formula:calc-manual-saved-values` confirms that Excel evaluates CF against the saved `<v>` on open. A volatile `TODAY()` under manual mode needs `formula:calc-manual-today`.
  - **One evaluation context (round-1 #5, split in round 2):** the document policy `CfDocumentEvaluation` (new) `{ calcMode, valuePolicy, dateSystem, timeZone }` is built **once per document** by `cfDocumentEvaluationOf(doc)` (resident, from `XlsxDocument.calc` and `doc.valuePolicy`) or `cfDocumentEvaluationOfPaged(workbook)`. Each preparation then takes **one clock snapshot** with `cfEvaluationContextFor(policy, now)`, giving `CfEvaluationContext = CfDocumentEvaluation & { clock: { epochNowMs, clockGen } }`. The policy is never re-sampled; the clock is never shared between preparations. **Omitted policy:** a document whose `valuePolicy` is undefined resolves through `workbookValuePolicy` (`value-policy.ts:11`, i.e. `'engine'`), and an absent `calc` resolves to `calcMode: 'auto'`, exactly as the workbook evaluator treats them; only a missing **context** (the route got none) withholds. The paged import records `calcMode` on `PagedWorkbook` beside `fullCalcOnLoad` (`pagedImport.ts:196–198`), and the worker forwards the context with its sheet metadata. Routes: `prepareCfPaint`, `prepareCfTextScope`, `readConditionalFormatting`/`prepareCfRemoteOnce`, `getPaintables` (`paint.ts:425`), `prepareSheetAssets` (`sheetAssets.ts:259`), and the a11y/copy/find scopes.
  - **Missing context withholds:** a route called without the context (a sheet-only direct API call) evaluates no formula rule: each is unresolved (`cf-evaluation-context-missing`). There is no default value policy for formula rules. The non-formula Stage 1–4 defaults (`render.ts:322` `'cached'`, `text.ts:438` `'engine'`) are left as they are, so no Stage 1–4 outcome changes.
  - **Cost if wrong:** manual-mode workbooks' formula rules stay plain.
- **Ruling S5-M: clock and volatile functions.**
  - **Spec tie:** §5.3 lines 144–147 and §5.11 line 211.
  - **`TODAY()`:** it reads the same clock `timePeriod` uses: the Stage 1 measured local-date policy (`timePeriod:*` published, `docs/validation/xlsx-cf.md:72`). It needs `formula:fn-today`.
  - **One clock per preparation:** the clock travels inside `CfEvaluationContext` (S5-L) into both `evaluatePredicate` and the formula context. The epoch is sampled once per preparation, and `clockGen` is the local day serial of that sample. That replaces the hardcoded `clockGen: 0` at `text.ts:435` and `viewerSession.ts:3329` for sheets with clock-dependent formula rules only. The interactive midnight tick (Ruling M2) then invalidates formula rules exactly as it does `timePeriod`.
  - **`NOW()`:** withheld (OD-14, decided 2026-11-16). Its sub-day value would need a periodic repaint the viewer does not have.
  - **Unaffected rules:** a formula without `TODAY` never reads the clock, so its stamp is clock-free.
  - **Cost if wrong:** `NOW`-based rules stay plain.
- **Ruling S5-N: work and memory for repeated aggregates.**
  - **Spec tie:** §6.3 line 245 and §7.
  - **Default (round-1 #2, #3):**
    - The snapshot is built once per preparation from the union of the **per-cell** footprints (hoisted ranges excluded, S5-HOIST), and is the only `SparseCellStore` the engine sees.
    - Work is charged where the engine reads: per `read`, per `stored()` element, plus the sort charge on entering `stored()` (Global Constraints). Repeated aggregates in one formula (`SUM(A:A)+SUM(A:A)`) and per-cell varying criteria (`COUNTIF($A:$A,A2)`) are therefore charged per traversal occurrence, not per footprint.
    - Evaluator scratch is reserved before each call from the payload-sensitive bound in the caps table; a refusal means the call never runs.
    - Each call's static work bound is ≤ 128 units, so calls never need to yield; the preparation loop yields and checks cancellation between cells (round-2 #2).
    - Absolute-range aggregates are hoisted to once-per-rule scalars or frequency maps by the statistics machinery (S5-HOIST); remaining cell-invariant subtrees under the per-call cap are memoised per preparation in a CF-owned map charged per entry.
    - A refusal of one rule's work or footprint never refuses another rule.
  - **Cost if wrong:** large formula sheets refuse with `cf-resource-limit`, a diagnostic, never a wrong format.
- **Ruling S5-HOIST: absolute-range aggregates are hoisted out of per-cell evaluation (orchestrator ruling).**
  - **Spec tie:** §6.3 line 245 ("an expression like `SUM(A:A)` cannot repeat a million-cell scan across visible cells") and §7 (cooperative checkpoints).
  - **What is hoisted.** At compile, `planCfHoist` (new) marks each `call` node whose reference arguments are all **absolute** (every axis `$`-pinned, or a whole column/row with pinned endpoints) and hoists it:
    - **Invariant aggregates:** `SUM`, `COUNT`, `COUNTA`, `AVERAGE`, `MIN`, `MAX` over an absolute range, and `COUNTIF`/`SUMIF` whose criterion is invariant (a literal or an absolute single-cell reference). Each becomes **one scalar per rule**.
    - **Per-cell criterion:** `COUNTIF(<absolute range>, <relative single-cell reference>)`, the duplicate-highlight idiom `COUNTIF($A:$A,A2)>1`, becomes **one value→count frequency map per rule**. The per-cell lookup then costs one unit.
  - **How it is computed (round-3 #3).** Through the existing chunked scanning infrastructure of the statistics loop (`service.ts:188–202`, `prepareStatistics` at `statistics.ts:793`, checkpoints every 128 units, fresh fallback), but with **engine-ordered reducers**, never the existing statistics totals: the statistics sum is Neumaier-compensated (`statistics.ts:527`) and gives 1 for `[1e16, 1, -1e16]` where the engine's sequential `SUM` gives 0. Each hoist reducer accumulates sequentially in reference order (areas in argument order, row-major within an area, as the engine's sparse cursor sorts them, `refs.ts:200`), and applies the engine's own rules for the first error in that order, overflow (`finiteOrNum`, `evaluator.ts:332`) and the empty range. The reducer asserts row-major arrival and refuses the hoist if a source yields out of order. A cross-sheet absolute range is scanned through the dependency's canonical store (S5-K).
  - **Exact semantics or refused.** The hoisted values must equal what the engine's own function returns on the same values:
    - scalar aggregates: equal to the engine's call on the same values, by the ordered reducer above; any value class whose engine treatment the reducer does not reproduce refuses the hoist;
    - the frequency map reproduces `COUNTIF`'s plain-equality arm only (`matches`, `functions/cond.ts:198`; criterion parsing in `criterionOf`, `:167–194`): a number criterion counts equal numbers; a text criterion counts strings that the engine's matcher (`matchWildcard`, `functions/wildcard.ts:85`) matches with no wildcard token. Its key is `cfCountifKey` (new), **not** `cfValueKey` (`statistics.ts:38`), which merges numeric text with numbers. Keys are computed, validated and folded **during the chunked hoist scan** and charged there (one unit per cell plus the key's string units); the per-cell lookup is then one unit.
    - **Refused per cell** (unresolved, never approximated), mirroring the engine's own refusals by calling or citing them rather than restating thresholds: a criterion that is text containing `*`, `?` or `~`; text with a comparison prefix (`criterionOf`'s prefix rule, `cond.ts:180–183`); an **error-literal criterion**, i.e. any text beginning with `#` (the engine turns `"#N/A"` into an error criterion that counts genuine errors, `cond.ts:189`); a criterion longer than the engine's criterion limit (`cond.ts:179`, `conditional-long-criterion-unverified`); a criterion whose matcher DP against the **longest string in the range** (recorded by the scan) would exceed the engine's DP policy (`DEFAULT_MAX_DP_CELLS` at `functions/wildcard.ts:28`, checked at `:137`, `conditional-wildcard-resource-limit`); non-ASCII text; **numeric-looking** text, defined by the engine's own rules: the numeric-criterion pattern of `criterionOf` (`cond.ts:185`) or `coerceToNumber` succeeding (`evaluator.ts:346`), so `" 1 "` counts; the empty string, a blank, a boolean or an error. To mirror without restating, Task 34 exports the two engine limits (`cond.ts:179`'s literal as `CONDITIONAL_MAX_CRITERION_LENGTH`, and `DEFAULT_MAX_DP_CELLS`) from the engine as named constants with no behaviour change, re-running the engine's conditional and wildcard tests.
    - **Refused for the whole rule:** a range that holds text that is numeric-looking by `coerceToNumber` (the engine reports `conditional-numeric-text-unverified` for number criteria, `cond.ts:227`), a boolean (`conditional-boolean-range-unverified`, `cond.ts:221`), or non-ASCII text (case-fold nuance).
    - Every other `COUNTIF`/`SUMIF` criterion shape (operators, wildcards, relative ranges) is not hoisted.
  - **The per-cell remainder.** The per-cell formula then evaluates against the hoisted scalars and map lookups under the 128-unit per-call cap; a hoisted node counts one unit.
  - **Cache key and invalidation.** Hoisted results are memoised per `(rule, normalised range, valueGen of the range's sheet, dependencyGen)` in the statistics session cache, and invalidated exactly like it: whole sheet, on any generation change that keeps the lease (Lesson 5).
  - **Charging.** Chunked, charged before allocation, with the frequency map's growth peak (old plus new table) reserved before it grows, and refused atomically: a refused or cancelled hoist publishes nothing and leaves nothing charged; the rule is `cf-resource-limit` (or cancelled).
  - **Budgets (round-3 #4).** Each hoist scan pays from its own statistics scan budget (`createCfWorkBudget({ signal, scan: true })`, as range statistics do at `service.ts:189`; on the paint path the factory at `render.ts:469`), and a retry pays from a fresh scan budget (`service.ts:199`). Neither counts against the per-cell preparation ceiling. The per-cell remainder pays from the per-call budget (`CF_LIMITS.activeUnits`) under the 128-unit per-call cap; the snapshot build and its fallback pay from the per-cell ceiling (2 × `activeUnits`). Generation rechecks stay: each hoist's source stamp is rechecked after its scan and before adoption (S5-K), and a change cancels it. **Whole columns** clamp to populated bounds exactly as Stage 1–4 statistics do (§6.3 line 243): a declared 1M-row `A:A` with sparse population scans its populated cells only, while a column with about a million **populated** rows refuses at the scan budget, as Stage 1–4 bars and scales already do (`docs/validation/xlsx-cf.md:1235`).
  - **Fallback.** A relative-range aggregate, or one that cannot be hoisted, keeps the per-cell path under the cap, or is refused.
  - **Gates:** `formula:hoist-aggregate` (scalar aggregates) and `formula:hoist-countif-frequency` (the map), each lifted by the `stage5-formula-hoist` oracle row, plus the function family capability.
  - **What it re-admits.** In the Task 31 matrix: `=COUNTIF($A:$A,A2)>1` and `=A2>AVERAGE($A$2:$A$9)` (`stage5-formula-functions`), `=MIN($A$2:$A$9)+1` (`stage5-cfvo-formulas`), and the four configurations of the new `stage5-formula-hoist` row. In the corpus, 0: `corpus/` holds no CF file (`docs/validation/xlsx-cf.md` §4). Task 39 reports the measured counts.
  - **Cost if wrong:** duplicate-highlight rules over text with wildcard characters, numeric text or non-ASCII letters stay plain.
- **Ruling S5-O: cross-kind composition.**
  - **Spec tie:** §5.7 and §10 item 5.
  - **Default:** a resolved formula rule is an ordinary DXF rule in `composeCell` (`compose.ts:300`). Its `stopIfTrue` halts like any DXF `stopIfTrue` (`halts`, `:79`). A formula bound changes only the visual's threshold, never its slot.
  - **Per-pair gates:** every cross-kind outcome with a formula rule on either side needs `formula:composition-<pair>`:
    - `expression-dxf`, `expression-scale`, `expression-bar`, `expression-icon`;
    - `stop-ahead-scale`, `stop-ahead-bar`, `stop-ahead-icon`;
    - `overlapping-expressions`;
    - `unresolved-expression-ahead-<kind>`.
  - **Existing guards stay:** the Stage 1 multi-positive guards (`service.ts:305–322`) and the Stage 2–4 slot gates also stay in force.
  - **Cost if wrong:** mixed sheets withhold their lower rules.
- **Ruling S5-STUB: every assertion that pins "withheld until Stage 5".**
  - **Spec tie:** §5.7 and §6.5.
  - **Exports unchanged:** the Stage 1–4 registries and exports stay byte-unchanged (`CF_STAGE1–4_*`, `CF_STAGE3_UNSUPPORTED`, `CF_STAGE4_UNSUPPORTED`, every `VERIFIED_STAGE1–4_*`).
  - **Inventory (round-1 #8; redone in round 2 by grepping every `<formula>`/`<xm:f>` text in `tests/xlsx-cf-*` and classifying its shape, not by comments) and successors.** Each row's successor is "Excel's recorded outcome under the named capabilities, or withheld"; nothing else is re-pointed:

  | Assertion | Cells | Successor capabilities |
  |---|---|---|
  | `STAGE5_OR_DEFERRED_WITHHELD` (`tests/xlsx-cf-oracle-conformance.test.ts:109–116`, asserted at `:225–226`) | `stage1-blanks-errors` D2 (`ISBLANK(D2)`), G6 (`NOT(ISBLANK(G2))`), J8 (`ISNA(J2)`), M6 (`NOT(ISNA(M2))`) | `formula:predicate-divergent`, `formula:fn-information`, `formula:fn-logical`, `formula:truth-boolean`, `formula:anchor-single` |
  | same set | `stage1-priority-composition` G2, G3 (`MOD(G2,2)=1`) | `formula:fn-math`, `formula:truth-boolean`, `formula:composition-stub-successor-dxf` |
  | `FINDINGS` (`:164–167`) | `stage2-scale-composition` L2–L5 | `formula:composition-stub-successor-scale` |
  | `STAGE3_NONBAR_WITHHELD` (`:171–172`) | `stage3-bar-composition` I2–I4 | `formula:composition-stub-successor-bar` |
  | `stage4StubWithheld` (`:283–286`, count 4 at `:366`) | `stage4-icon-composition` L cells | `formula:composition-stub-successor-icon` |
  | Parser and predicate unit tests pinning formula shapes as unresolved: `tests/xlsx-cf-parse.test.ts:153–160` (non-literal, `expression`), `:226` (x14 divergent), `tests/xlsx-cf-predicates.test.ts:183` (non-literal B1), `:264–266` (divergent `SEARCH`, stays unresolved), `tests/xlsx-cf-scale-parse.test.ts:216–242`, `tests/xlsx-cf-bar-parse.test.ts:331–332`, `tests/xlsx-cf-icon-parse.test.ts:234–237` | synthetic | each is pinned under `withExactCfEvidence([])` (or keeps an unregistered function), so it asserts the same refusal without evidence and cannot flip silently when Stage 5 publishes; a twin under the injected capabilities asserts the resolved shape |
  | Cap and lifecycle tests that use `type="expression"` as a generic stub: `tests/xlsx-cf-parse.test.ts:674,771,821,828,968,995,1109,1135,1180,1201`, `tests/xlsx-cf-paged-parse.test.ts:101` | synthetic | unchanged where the stub has no formula (still `cf-invalid-rule`/unresolved); where it carries `<formula>A1</formula>` the test pins `withExactCfEvidence([])` |
  | R1c timePeriod tests (`tests/xlsx-cf-parse.test.ts:257–262`, `tests/xlsx-cf-paint-fixes.test.ts:165`) | synthetic | rewritten per S5-G (Task 33), together with the ten-period table's `lastWeek` row (`tests/xlsx-cf-parse.test.ts:240`) |
  | `F6` (`tests/xlsx-cf-parse.test.ts:848–852`): `cellIs` operand `"a"&"b"` (round-2 #5) | synthetic | pinned under `withExactCfEvidence([])` (still unresolved); twin under `formula:cellis-operand-expression` + dialect/anchor resolves to a formula operand |
  | `F7` (`tests/xlsx-cf-parse.test.ts:867–871`): off-anchor `containsBlanks LEN(TRIM(Z99))=0` (round-2 #5) | synthetic | pinned under `withExactCfEvidence([])`; twin under `formula:predicate-divergent` compiles as an expression reading Z99 |
  | `F4` (`tests/xlsx-cf-parse.test.ts:1024–1035`): pinned `LEN(TRIM($A$1))=0`, `LEN(TRIM(A$1))=0` | synthetic | pinned under `withExactCfEvidence([])`; twin compiles as divergent expressions |
  | Divergent blank/error parse rows (`tests/xlsx-cf-parse.test.ts:180–181`, `ISBLANK(G2)`, `ISNA(H2)`) and `cellIs` operand `B2>100` (`:119`) | synthetic | pinned under `withExactCfEvidence([])`; twins under complete evidence |
  | `tests/xlsx-cf-paged-semantics.test.tsx:74`: `LEN(A1)>100000` used to make a rule unresolved | synthetic | switched to an unregistered function (`ISEVEN(A1)`) so it stays unresolved under any published evidence |
  | Parity tests with expression shapes (`tests/xlsx-cf-paged-parse.test.ts:61,203`) | synthetic | unchanged: they compare the two paths, which stay deep-equal |
  | `SEARCH`-based rows (`tests/xlsx-cf-parse.test.ts:175–177,198–199,217–218,873,883,1039,1050`; `tests/xlsx-cf-predicates.test.ts:264`; `tests/xlsx-cf-oracle.test.ts:1870,4026`) | synthetic | unchanged: canonical rows keep the Stage 1 path, divergent ones stay unresolved (`SEARCH` unregistered) |

  - **Reused captures:** the four composition cases' committed captures are re-read through `CF_STAGE5_REUSED_CAPTURES`; `stage1-blanks-errors` is reused the same way.
  - **One commit:** the conformance re-pointing lands in one named commit; a grep test asserts it touches only the rows above.
  - **Cost if wrong:** none beyond today. Without the capabilities those cells keep withholding.
- **Ruling S5-P: every view agrees.**
  - **Spec tie:** §6.4 and §6.7.
  - **Inputs per leg:** each leg takes its inputs from one place:
    - values: its `CfValueSource` plus `CfWorkbookSource`, both over canonical overlay-aware stores (S5-K);
    - calculation mode, value policy and clock: one `CfEvaluationContext` (S5-L);
    - a route lacking either withholds formula rules identically on paint, a11y and copy/find.
  - **Table:** Task 38 pins this as a source table (static `render.ts:287`; print and export through `paint.ts:425`; resident `sheetAssets.ts:259`; text `text.ts:364`; worker `prepareCfRemoteOnce` `pagedSource.ts:1773`; forced-paged the same).
  - **Missing inputs:** a leg without `CfWorkbookSource` (a direct `renderSheet` caller with one sheet) withholds cross-sheet rules identically on paint, a11y and copy/find.
  - **Blank synthesis:** `canRuleMatchBlanks` returns true for every resolved formula-bearing kind (`expression`, formula-operand `cellIs`, divergent predicates), so blank synthesis stays on wherever a formula can match a blank.
  - **Cost if wrong:** a leg that cannot see another sheet stays plain where others paint. That is a parity failure the gate must catch, so the tests compare legs pairwise.
- **Ruling S5-Q: accessibility, copy and find.**
  - **Spec tie:** §6.7.
  - **Default:** a formula rule's descriptor is the Stage 1 DXF descriptor (`describeCfWin`, `describe.ts:65`), derived from the winning properties, never from formula text (Lesson 10). Copy and find follow the effective number format exactly as Stage 1 does.
  - **Cost if wrong:** none new.
- **Ruling S5-R: wire and bytes.**
  - **Spec tie:** §6.8.
  - **Default:** no per-cell wire field is added. Formula outcomes compose into the existing `properties`/`scale`/`bar`/`icon` fields of `CfCellStateWire` (`contracts.ts:396`).
  - **New reply data:** a reply grows only by `CfReadResult.dependencies`, at most 16 entries, priced by `measurePayloadBytes`.
  - **No shared definitions:** compiled formulas never cross the wire. The worker compiles from its own rules, and the main thread consumes no per-rule formula data (the S4-Q pattern).
  - **Cost if wrong:** reply headroom only.

### Deferral sweep

Built from the code (`finishUnresolved` sites in `parse.ts`, `unmeasured()` sites in `predicates.ts`, `CF_STAGE3_UNSUPPORTED`/`CF_STAGE4_UNSUPPORTED`), not only from the plan's "until Stage 5" text.

| Deferred shape | Where on main | Disposition |
|---|---|---|
| `expression` rule (base and x14) | `parse.ts:1526–1527` | Picked up: S5-B–E, Task 35 |
| Non-literal `cellIs` operand | `parse.ts:1449–1451` | Picked up: S5-F, Task 35 |
| `cellIs` text operators | `parse.ts:1439–1440` | **Re-deferred** (S5-F): Excel's dialog writes the standalone types; recorded only |
| Divergent text/blank/error formula | `parse.ts:1470,1482` | Picked up: S5-G, Task 35 (the `SEARCH` forms stay refused under S5-B) |
| Divergent `timePeriod` formula | none: main resolves it from the attribute (R1c, `parse.ts:1490`); `:1489` is a namespace check | Fail-open fixed: classified on both paths, divergent unresolved until measured (S5-G, Task 33) |
| Formula on `duplicateValues`/`uniqueValues` | `parse.ts:1511–1512` | **Re-deferred** (S5-G): Excel writes none |
| Scale `type="formula"` / reference `@val` | `parse.ts:807–809,834–842` | Picked up: S5-H, Task 37 |
| Bar `type="formula"` / reference `xm:f` | `parse.ts:982–983`, `cfvoLiteral` `:960–969` | Picked up: S5-H, Task 37 |
| Icon `type="formula"` / reference `xm:f` | `parse.ts:1225–1226` | Picked up: S5-H, Task 37 |
| Standalone x14 colour scale (OD-3) | `parse.ts:767–769` | Picked up conditionally: S5-I, Task 37 |
| `activePresent="1"` on scale/bar/icon | `parse.ts:772–774,1030–1032,1250–1251` | Conditional: S5-J, Task 37; expected to re-defer permanently |
| Multi-range formula anchor | `parse.ts:1419` | Picked up behind `formula:anchor-multi-range`, Task 34 |
| Stub compositions (`bar:`/`icon:composition-unresolved-stub`) | `cases.ts:4706`, `CF_STAGE4_UNSUPPORTED` | Picked up as `formula:composition-stub-successor-*` (S5-STUB), Task 35 |
| Cross-sheet values | none (single-sheet sources) | Picked up: S5-K, Task 36 |
| Clock / midnight for formulas | `text.ts:435`, `viewerSession.ts:3329` | Picked up: S5-M, Task 38 |
| Cached/manual policy | `render.ts:322` vs `text.ts:438` | Picked up: S5-L, Task 38 |
| `NOW()` | `functions/datetime.ts:617` | **Re-deferred** (OD-14): needs periodic repaint |
| Defined names, external refs, tables | spec §2 | **Out of scope**: refused (S5-B) |
| Lookup and dynamic-array functions | `functions/lookup.ts:316` | **Re-deferred** (OD-13, decided) |
| `bar:pairing-standalone` | `CF_STAGE3_UNSUPPORTED` | Not Stage 5 (Excel resaves it paired); unchanged |

### Planned files and interfaces

| File | Stage 5 responsibility |
|---|---|
| `scripts/cf-refs/cases.ts`, `generate.ts`, `capture.ts`, `verify.ts`, `build.ts`, `audit.ts`, `scripts/cf-compare.ts` | `CF_STAGE5_*`; formula markers; semantic DisplayFormat route; Stage 5 audit and `--publish-stage5`; reused Stage 1–4 captures |
| `tests/fixtures/xlsx-cf/excel-authored-formulas.xlsx` (new, orchestrator) | Excel-dialog-authored formula, cfvo-formula and x14 XML (Lesson 12) |
| `src/xlsx/cf/formula.ts` (new) | Compile and admit, allow-list, translator, footprint, truth readings, per-cell evaluation |
| `src/xlsx/cf/dependencies.ts` (new) | Dependency snapshot, `SparseCellStore` over it, `CfWorkbookSource` adapters, pending and taint |
| `src/xlsx/cf/model.ts`, `parse.ts`, `evidence.ts` | Formula definition arms; parse wiring on both paths; `formulaRequiredCapabilities` |
| `src/xlsx/cf/service.ts`, `predicates.ts`, `colorScale.ts`, `dataBar.ts`, `iconSet.ts`, `compose.ts`, `limits.ts` | Formula arms, cfvo formula thresholds, caps |
| `src/xlsx/cf/cache.ts`, `render.ts`, `text.ts`; `src/xlsx/view/pagedSource.ts`, `contracts.ts`; `src/worker/viewerRemoteCf.ts`, `viewerRemoteSource.ts`, `viewerSession.ts`; `src/components/xlsx/InteractiveSheetDoc.tsx`; `src/xlsx/cfLoader.ts` | `dependencyGen`, dependencies on the wire, cross-sheet invalidation, clock and policy threading, blank-synthesis arms |

**Consumes** (all printed on `fc61c6c6`):
- **Engine:** `parseFormula` (`parser.ts:859`), `tokenize` (`lexer.ts:149`), `AstNode` (`types.ts:206`), `evaluateNode` (`evaluator.ts:717`; `evaluateFormulaInternal` `:1100` is deliberately not consumed), `chargeMaterialization` (`:179`), `markUnsupported` (`:499`), `isEvaluationError`/`formulaError` (`:109,105`), `EvaluationContext`/`MaterializationBudget`/`ResolvedSemantics`/`CellAddress`/`ReferenceServices` (`types.ts:92,90,279,234,381`), `createReferenceServices`/`ReferenceServiceDeps`/`SparseCellStore` (`refs.ts:810,50,43`), `FUNCTIONS` (`functions.ts:508`), `epochToSerial` (`serial.ts:328`).
- **CF model and evaluation:** `CfDefinition`/`CfRuleOrigin`/`CfLiteral`/`CfStamp`/`CfOutcome`/`CfValueSource` (`model.ts:55,14,10,5,6,98`), `cfvoLiteral` (`parse.ts:960`), `readFormulas` (`:1556`), `createCfService` (`service.ts:71`), `estimateCfPrepareBytes` (`:37`), `evaluatePredicate` (`predicates.ts:461`), `CfDateContext` (`:29`), `composeCellAsync` (`compose.ts:365`), `createCfWorkBudget`/`createCfMemory`/`CF_LIMITS` (`limits.ts:276,70,43`), `sameCfStamp` (`cache.ts:5`), `hasEvidence`/`withExactCfEvidence` (`evidence.ts:88,28`).
- **Sources and views:** `residentSource`/`storeCfSource`/`prepareCfPaint` (`render.ts:68,129,287`), `prepareCfTextScope` (`text.ts:364`), `readConditionalFormatting`/`prepareCfRemoteOnce`/`cfRulesVersion` (`pagedSource.ts:897,1773,1511`), `CfReadResult`/`CfCellStateWire`/`measureCfCellBytes` (`contracts.ts:422,396,483`), `RemoteCfCache` (`viewerRemoteCf.ts:147`), `isValueUnavailable` (`view/values.ts:50`), `measurePayloadBytes` (`view/payloadBytes.ts:62`), `canRuleMatchBlanks` (`cfLoader.ts:100`).
- **Oracle:** `buildColorProbeScript`/`buildDecorationProbeScript`/`probeLiveExcelCellObservations` (`capture.ts:213,192,587`), `captureUnderLockGuard` (`:493`), `replicationShortfall` (`verify.ts:5061`), `verifyCfCase`, `auditStage4Manifest`/`stage4Publication`/`renderVerifiedCapabilitiesSource` (`audit.ts:665,825,924`), `CF_CASES` (`cases.ts:5604`).

**Produces** (every name below was grepped absent on `fc61c6c6`; the creating task is given):
```ts
// Task 31 — scripts/cf-refs
CfPartMarker.formula?: { kind: 'expression' | 'operand' | 'cfvo' | 'activation'; anchor?: string; dialect: 'base' | 'x14' }
CfPartMarker.scale.cfvo[].f?: string ; CfPartMarker.scale.dialect?: 'base' | 'x14'          // widened in place
CfPartMarker.activePresent?: boolean
CF_STAGE5_CASES / CF_STAGE5_CAPABILITIES / CF_STAGE5_NONDISCRIMINATING / CF_STAGE5_UNSUPPORTED   // cases.ts
CF_STAGE5_REUSED_CAPTURES: readonly { caseId: string; refs: readonly string[]; capabilityId: string }[]
auditStage5Manifest(manifest: CfManifest, options?: { rootDir?: string }): Stage5AuditReport     // audit.ts
Stage5AuditReport (new): { ok: boolean; stageComplete: boolean; blocked: readonly string[]; perCapability: readonly Stage5CapabilityAudit[] }
RenderVerifiedOptions.publishStage5?: boolean                                                    // --publish-stage5
// Task 32 — src/xlsx/cf/formula.ts, model.ts, parse.ts
type CfFormula = Readonly<{ ast: AstNode; nodes: number; atoms: readonly CfFormulaRef[]; invariant: boolean; usesClock: boolean; families: readonly string[] }>
type CfFormulaRef = Readonly<{ sheet?: string; firstRow: number; lastRow: number; firstCol: number; lastCol: number; absRow: boolean; absCol: boolean; axis?: 'col' | 'row' }>
CF_FORMULA_FUNCTIONS: ReadonlyMap<string, CfFormulaFunction>   // frozen: NAME → family and argument profile
compileCfFormula(text: string, anchor: { row: number; col: number } | undefined, memory: CfMemory):
  CfLease<CfFormula> | { refused: 'cf-unsupported-rule' | 'cf-resource-limit'; reason: string }   // reserves before parseFormula
type CfFormulaUse = 'expression' | 'operand' | 'cfvo' | 'activation'
formulaRequiredCapabilities(f: CfFormula, use: CfFormulaUse, rule: Readonly<{ dialect: 'base' | 'x14'; rangeCount: number; sourceType: CfRuleType }>):
  Readonly<{ parse: ReadonlySet<string>; evaluation: ReadonlySet<string> }>
canonicalTimePeriodFormula(period: CfTimePeriod, formulas: readonly string[], anchor: string | undefined): boolean   // Task 33
// CfDefinition gains: { kind: 'expression'; formula: CfFormula; sourceType: CfRuleType }
//   and cellIs gains operands: readonly (CfLiteral | { formula: CfFormula })[]
// CfScaleBound / CfBarBound / CfIconThreshold.bound gain { type: 'num' | 'percent' | 'percentile' | 'formula'; formula: CfFormula }
CfRuleOrigin.dependencySheets?: readonly string[]          // referenced sheet names, data only
// Task 34 — src/xlsx/cf/formula.ts, dependencies.ts, limits.ts
type CfEdgeMode = 'wrap' | 'ref-error'
translateCfFormula(f: CfFormula, dRow: number, dCol: number, edge: CfEdgeMode, memory: CfMemory): CfOutcome<CfLease<AstNode>>
cfFormulaFootprint(f: CfFormula, targets: readonly SourceRect[], anchor: { row: number; col: number }, edge: CfEdgeMode): readonly { sheet?: string; rect: SourceRect }[] | 'over-cap'
type CfTruthReading = 'number-nonzero-true' | 'boolean' | 'text-no-match' | 'text-true' | 'blank-no-match'
  | 'error-no-match' | 'error-halts' | 'range-implicit' | 'range-top-left' | 'range-no-match'
type CfResult = Readonly<{ kind: 'number' | 'boolean' | 'text' | 'blank' | 'error'; value: ElementValue } | { kind: 'range'; matrix: MatrixValue; area: SourceRect }>
cfTruth(result: CfResult, at: { row: number; col: number }, readings: ReadonlySet<CfTruthReading>): CfOutcome<boolean | 'halt'>
cfLiteralToElement(v: CfLiteral): ElementValue ; cfElementToLiteral(v: EvaluationValue): CfLiteral | undefined
interface CfDependencySnapshot { read(sheet: string | undefined, row: number, col: number): ElementValue; readonly tainted: boolean; reset(): void; dispose(): void
  viewFor(effective: readonly { sheet?: string; rect: SourceRect }[], budget: CfWorkBudget):   // effective geometry; refuses if not fully covered, before traversal CfOutcome<{ store: SparseCellStore; metrics: CfPayloadMetrics }> }   // the per-call store
prepareCfDependencies(atoms: readonly { sheet?: string; rect: SourceRect }[], host: CfValueSource, workbook: CfWorkbookSource | undefined,
  budget: CfWorkBudget, memory: CfMemory, signal: AbortSignal): Promise<CfOutcome<CfLease<CfDependencySnapshot>>>
evaluateCfFormula(f: CfFormula, at: { row: number; col: number }, snapshot: CfDependencySnapshot, context: CfFormulaContext): CfOutcome<CfResult>
  // never via evaluateFormulaInternal; discards on any recorded feature, issue, cycle, refusal or taint (S5-E)
type CfFormulaContext = Readonly<{ sheetName: string; evaluation: CfEvaluationContext; semantics: ResolvedSemantics; edge: CfEdgeMode | undefined;
  materialization: MaterializationBudget; budget: CfWorkBudget; memory: CfMemory; invariants: CfInvariantMemo }>
type CfInvariantMemo (new): a charged per-preparation map keyed by compiled subtree and measured scope
type CfPayloadMetrics (new): Readonly<{ viewCells: number   // cardinality of the call's restricted store (stored cells in its translated atoms), an upper bound on every stored() enumeration
  atoms: readonly Readonly<{ cells: number; stringUnits: number; longestString: number }>[] }>   // measured while the snapshot is built
cfEvaluatorScratchBytes(f: CfFormula, payload: CfPayloadMetrics): number   // the bound reserved before each call
cfCallWorkUnits(f: CfFormula, payload: CfPayloadMetrics, criterionLengths: readonly number[]): number | 'unbounded'   // ≤ 128 or refused
CF_DEPENDENCY_CELL_BYTES: number ; CF_FORMULA_NODE_BYTES: number ; CF_SLOT_BYTES: number   // measurer-derived
CF_LIMITS.dependencyCells / formulaAtoms / referencedSheets / materializationCells   // new keys
// Task 36 — dependencies.ts, model.ts, contracts.ts, worker
interface CfWorkbookSource { sheetIdOf(name: string): string | undefined; store(sheetId: string): SourceStore | undefined }   // values and stamp from the same store
CfStamp.dependencyGen?: number                             // absent ≡ 0; sameCfStamp compares it
CfReadResult.dependencies?: readonly { sheetId: string; valueGen: number }[]
// Task 32 (contract and route wiring, before any formula evaluates) — formula.ts, render.ts, text.ts, pagedSource.ts, pagedImport.ts
type CfDocumentEvaluation = Readonly<{ calcMode: 'auto' | 'manual' | 'autoNoTable'; valuePolicy: ValuePolicy; dateSystem: '1900' | '1904'; timeZone: string }>
type CfEvaluationContext = CfDocumentEvaluation & Readonly<{ clock: Readonly<{ epochNowMs: number; clockGen: number }> }>
cfDocumentEvaluationOf(doc: Pick<XlsxDocument, 'calc' | 'valuePolicy' | 'semantics'>): CfDocumentEvaluation   // once per document
cfDocumentEvaluationOfPaged(workbook: PagedWorkbook): CfDocumentEvaluation
cfEvaluationContextFor(policy: CfDocumentEvaluation, now: () => number): CfEvaluationContext                 // once per preparation
PagedWorkbook.calcMode?: 'auto' | 'manual' | 'autoNoTable'           // recorded beside fullCalcOnLoad
// prepareCfPaint / prepareCfTextScope / CfRemotePrepareOptions gain `evaluation?: CfDocumentEvaluation`; absent ⇒ formula rules unresolved
CfFormulaFunction (new): Readonly<{ family: string; args: readonly ('scalar' | 'range')[]; text?: true
  geometry: readonly ('literal' | { resizedTo: number })[] }>   // effective read geometry per argument (S5-SUMIF-SHAPE)   // CF_FORMULA_FUNCTIONS values
CF_CHAR_SLOT_BYTES: number                                                                                 // measurer-derived
// Task 34 — S5-HOIST (formula.ts, service.ts, statistics.ts)
type CfHoist = Readonly<{ node: AstNode; kind: 'scalar' | 'countif-frequency'; fn: string; range: Readonly<{ sheet?: string; rect: SourceRect }>; criterion?: AstNode }>
planCfHoist(f: CfFormula): readonly CfHoist[]
cfCountifKey(value: CfLiteral): string | 'refuse'                       // plain-equality COUNTIF key; never cfValueKey
type CfHoisted = Readonly<{ scalars: ReadonlyMap<AstNode, ElementValue>; frequencies: ReadonlyMap<AstNode, ReadonlyMap<string, number>> }>
CfStatisticsSession.prepareHoists(rule: CfRule, hoists: readonly CfHoist[], stamp: CfStamp, source: CfValueSource, budget: CfWorkBudget,
  options?: CfStatisticsOptions): Promise<CfOutcome<CfLease<CfHoisted>>>
// Task 31 — capture.ts
buildManualObservationScript(workbookName: string, sheetName: string, refs: readonly string[]): string[]   // no `calculate` before the observation
CfOracleCase.calcObservation?: 'pre-calculate'                                                       // build/capture skip the unconditional calculate
```

### Execution, assignments and review

Implementers are Claude subagents: Sonnet for mechanical tasks, Opus for judgement-heavy ones. A different model family reviews every task: opencode, or codex when opencode is unavailable. Before dispatch, record the actual model of each seat, and never relabel a same-family review as independent. Implementers do not run Excel. The orchestrator executes every host-Excel step (AppleScript and System Events) while the owner keeps the display unlocked. Reports and reviews go in `.superpowers/sdd/<plan>/task-N-report.md` / `task-N-review.md`. After every agent, check `git status` and `git diff --stat` for test files. Task 39 ends with a high-effort whole-branch review by the other family.

| Task | Implementer | Reviewer |
|---|---|---|
| 31 Stage 5 Excel oracle, semantic route, Stage 5 audit, withheld-only conformance | Claude Opus (orchestrator executes host Excel) | opencode / codex |
| 32 Formula compile, admission, allow-list, model and both-path parsing | Claude Sonnet | opencode / codex |
| 33 Canonical/divergent `timePeriod` formula classification on both parser paths, R1c reconciliation | Claude Sonnet | opencode / codex |
| 34 Translator, truth readings, same-sheet dependency snapshot, budgets | Claude Opus | opencode / codex |
| 35 Formula DXF rules: expression, operands, divergent predicates, composition, stub successors | Claude Opus | opencode / codex |
| 36 Cross-sheet values, pending pages, `dependencyGen`, remote invalidation | Claude Opus | opencode / codex |
| 37 Formula cfvo bounds, x14 colour scales, `activePresent` | Claude Opus | opencode / codex |
| 38 Calculation policy, clock, every-view agreement, blank synthesis, a11y/copy/find | Claude Sonnet | opencode / codex |
| 39 Stage 5 validation, five-family closure gates, resource audit, PR evidence | Claude Sonnet | opencode / codex, plus whole-branch high-effort review |

**One PR (OD-11, decided 2026-11-16).** Stage 5 ships as one PR. Tasks run in order 31–39; Task 37 follows Task 36 because S2-K expects x14 scales to be cross-sheet. Task 39 publishes once, when the whole Stage 5 audit reports `stageComplete=true`.

### Review Focus (Stage 5)

1. A read outside the snapshot, on a pending page or of an unavailable cached value that yields a guessed `0`/blank instead of unresolved. Every taint must discard the result (S5-E, Tasks 34, 36).
2. An allow-list checked at run time, so `IF(FALSE, INDIRECT(…), 1)` passes, or a later `registerFormulaFunction` widens it. Admission must be static on the AST (S5-B, Task 32).
3. A translated reference past the sheet edge, or a multi-range anchor, that paints before its reading is measured (S5-C, Tasks 31, 34).
4. A cross-sheet commit after which the old batch still paints. `at()` must go stale synchronously through `dependencyGen`, on resident and worker legs alike (S5-K, Task 36).
5. A hoisted `COUNTIF($A:$A,A2)` frequency map that counts differently from the engine's `COUNTIF` (wildcards, numeric text, case, booleans), or a hoist that publishes after cancellation or is not invalidated with its range's generation (S5-HOIST, Task 34). Also `SUM(A:A)+SUM(A:A)` on a relative-range path, `LEN(CONCAT(A1:A64))` over 32,767-character strings and `IFERROR(COUNTIF(A:A,TRUE),TRUE)`: a per-call bound over 128 units or over the text cap refused before invocation, cancellation honoured between cells, every traversal occurrence charged before it happens, scratch reserved before the call, refusal before excess work, and any recorded unsupported feature discarding the result (S5-E, S5-N, Lessons 9, 15, Task 34).
6. A blank or range result reaching a number gate (blank→0 or engine projection before the CF shape gate), or two readings that predict the same observable published as distinct (S5-D, Task 34).
7. A view without the workbook source or with a different clock or policy that paints where another withholds (static vs text defaults `render.ts:322`/`text.ts:438`) (S5-L/M/P, Task 38).
8. A resolved `expression` that disables blank synthesis through `canRuleMatchBlanks`, or is priced at zero by `estimateCfPrepareBytes` (S5-P, Tasks 32, 38).
9. A "withheld until Stage 5" assertion re-pointed outside the S5-STUB inventory, a unit test that flips silently at publication, an unevidenced divergent `timePeriod` formula reaching the attribute predicate, or any Stage 1–4 export byte changed (S5-STUB, S5-G, Tasks 33, 35, 39).
10. A capability lifted from an evaluation rule changed after seeing the capture with no independent justification (post-hoc fitting), or a `[measure]` row with one decisive cell (Lessons 3, 14, Task 31).

### Task 31: Stage 5 Excel oracle, semantic route, audit, conformance (Claude Opus; review per table)

**Dependencies:** Stage 4 merged (PR #79, `fc61c6c6`). Host Excel runs only through the orchestrator. Preparing scripts is not evidence: every case without captures stays `recipe-pending`, and refused.

**Files:**
- Modify `scripts/cf-refs/cases.ts`: add the `CF_STAGE5_*` exports and `CF_STAGE5_GROUPS`/`CF_STAGE5_REUSED_CAPTURES` beside the Stage 4 exports; widen `CfPartMarker` as listed; extend `CF_CASES`/`CF_CAPABILITIES` (`:5604`).
- Modify `generate.ts` (templates diffed against the Excel-authored fixture; one malformed shape per workbook), `capture.ts` (DisplayFormat probes for the Stage 5 cases, the `dynamicRefresh` edit of a referenced sheet), `verify.ts` (formula marker semantics, `CfPartMarker.formula`), `audit.ts` (the guard at `:2207` accepts 5, `auditStage5Manifest` follows `:665`, `--publish-stage5` mirrors S3-WRITE and `--publish-stage4`), `scripts/cf-compare.ts` (Stage 5 cases).
- Modify `tests/xlsx-cf-oracle.test.ts`, `tests/xlsx-cf-oracle-conformance.test.ts`, `tests/fixtures/xlsx-cf/manifest.json`, plus case files and `refs/*`. Create `tests/fixtures/xlsx-cf/excel-authored-formulas.xlsx` (orchestrator) and `docs/validation/xlsx-cf-recipes-stage5.md`. Update `docs/validation/xlsx-cf.md`.

**Interfaces:**
- **Consumes:** `CfOracleCase`, `CfPartMarker`, `CfCellExpectation` (`fillSource`, `dynamicRefresh`), `verifyCfCase`, `buildColorProbeScript`, `buildDecorationProbeScript`, `probeLiveExcelCellObservations`, `captureUnderLockGuard`, `replicationShortfall`, `auditStage4Manifest`/`stage4Publication` (patterns), `renderVerifiedCapabilitiesSource`, `DEFAULT_CLOCK_ANCHOR_DATE` (`cases.ts:399`) and `--refresh-clock-date` (`build.ts:247`).
- **Produces:** the Task 31 block of **Produces**.

**Required Stage 5 oracle matrix.** Every row is `[measure]`. No expected outcome is written by the plan or by the code's authors. Each row has a same-content unformatted control block, positive and negative controls, a repair-free open, and retained markers diffed against Excel's resaved XML. Every configuration has ≥ 2 decisive cells (S4-REPLICATE), with the hypotheses fixed in the case before capture (Lesson 14).

| Case ID | Input / question (competing hypotheses) | Capabilities |
|---|---|---|
| `stage5-formula-authored` | Orchestrator authors through Excel's dialog: expression `=$B2>$C2` over B2:D9; `=MOD(ROW(),2)=0`; `cellIs > =$F$1`; `between =$F$1,=$F$2`; a scale with `num` `=$F$1` and `formula` bounds; a bar and an icon set with formula bounds; `=Sheet2!$A$1>0`; a scale bounded by `Sheet2!$A$1`. Output `excel-authored-formulas.xlsx`: the template source. | `formula:dialect-base`, `formula:dialect-x14` |
| `stage5-formula-truthiness` | Expression results 1, 0, −0.5, TRUE, FALSE, `"abc"`, `""`, `="TRUE"`, a blank reference, `#N/A`, `#DIV/0!`, `=A2:A4` (target row 3, so implicit intersection ≠ top-left), each over ≥ 2 cells, each once with a lower matching fill and once with `stopIfTrue="1"` ahead of a lower matching fill. Nested shapes `=IF(TRUE,A2:A4,0)` and `=IFERROR(A2:A4,0)` at target row 3 with A2 ≠ A3 record `formula:truth-range-nested` (round-2 #4). Hypotheses per S5-D, stated as observables: format/no format × lower rule shown/halted. | `formula:truth-number`, `-boolean`, `-text`, `-blank`, `-error`, `-range`, `-range-nested` |
| `stage5-formula-anchor` | `=A2>0` over B2:B9; `=$A2`, `=A$2`, `=$A$2`; multi-range `D5:D8 B2:B4` with values making h1/h2/h3 differ at ≥ 2 cells each; Excel's resaved `sqref` order recorded. | `formula:anchor-single`, `-mixed-abs`, `-multi-range` |
| `stage5-formula-edge` | Rows: `sqref="A1:A4"` with `=A1048575>0` (offset +1048574): A3 and A4 overflow, so two targets decide w1 (wrap to A1, A2) vs w2 (`#REF!`), with A1, A2, A1048575, A1048576 holding values that separate them. Columns: `sqref="A1:D1"` with `=XFC1>0`: C1 and D1 overflow. Each configuration has two overflowing targets. | `formula:relative-wrap` (rows and columns as separate configurations) |
| `stage5-formula-functions` | One configuration per family in `CF_FORMULA_FUNCTIONS`, plus `=A2=""` and `=ISBLANK(A2)` over an unstored blank and `""` (the blank-versus-empty-text outcome comes from Excel, round-1 #12): `=AND($A2>0,$B2<5)`, `=ISNUMBER(A2)`, `=LEN(A2)>3`, `=MOD(A2,3)=0`, `=A2>AVERAGE($A$2:$A$9)`, `=COUNTIF($A:$A,A2)>1`, `=YEAR(A2)=2026`, `=ROW()=COLUMN()`; inputs include cached formula cells and errors. | `formula:fn-logical`, `-information`, `-text`, `-math`, `-aggregate`, `-conditional-aggregate`, `-date`, `-row-column` |
| `stage5-formula-hoist` | **Admitted configurations** (each ≥ 2 decisive cells): `=COUNTIF($A:$A,A2)>1` over A2:A40 holding numbers, ASCII text in mixed case (`"abc"`/`"ABC"`), blanks and errors, with duplicates and singletons; `=A2>AVERAGE($A:$A)`; `=A2=MAX($A:$A)`; `=COUNTIF($A:$A,"x")>2`. **Refused configurations**, recorded only, each in its own column and never evidence for success: a range with numeric text (`"1"`, `" 1 "`), a range with a boolean, non-ASCII text, wildcard criteria (`"a*"`), an `"#N/A"` criterion over a range holding `#N/A`. | `formula:hoist-aggregate`, `formula:hoist-countif-frequency` |
| `stage5-cellis-operands` | `cellIs >` with `$F$1`, `F2` (relative), `$F$1*2`; `between` with reversed referenced bounds; operands evaluating to `"100"`, blank, `#N/A`, TRUE. | `formula:cellis-operand-reference`, `-relative`, `-expression`, `-text`, `-blank`, `-error` |
| `stage5-predicate-formulas` | Divergent `containsBlanks ISBLANK(D2)`, `notContainsErrors NOT(ISNA(E2))`, `timePeriod` today with `<formula>A2=TODAY()-3</formula>` (attribute vs formula); a hand-authored `cellIs containsText` (record only). | `formula:predicate-divergent`, `formula:predicate-timeperiod-divergent` (the divergent `timePeriod` cells: ≥ 2 where attribute and formula disagree in each direction) |
| `stage5-clock-today` | `=A2=TODAY()` with D−2..D+2 and fractional serials under the Stage 1 anchor (`--refresh-clock-date`); `=A2<NOW()` recorded only. | `formula:fn-today` |
| `stage5-calc-manual` | `<calcPr calcMode="manual"/>` with formula cells whose saved `<v>` differs from a recompute; expressions reading them; `TODAY()` in CF. **Capture route (round-1 #4):** `calcObservation: 'pre-calculate'`. Observed with `buildManualObservationScript` (application calculation mode, values and DisplayFormat) before any `calculate`, before resave, with no other workbook open; then one explicit `calculate` and a second observation, recorded separately. Hypotheses: CF follows saved values vs recomputed. | `formula:calc-manual-saved-values`, `formula:calc-manual-today` |
| `stage5-composition` | Expression p1 `stopIfTrue` true/false ahead of a DXF, a scale, a bar and an icon; an expression fill vs a scale fill both orders; two overlapping expressions; an unresolved expression (`INDIRECT`) p1 ahead of each kind. Plus the reused Stage 1–4 stub captures (`CF_STAGE5_REUSED_CAPTURES`). | `formula:composition-*`, `formula:composition-stub-successor-dxf/-scale/-bar/-icon` |
| `stage5-cross-sheet` | `=Sheet2!A2>0` (relative cross-sheet); `=$A2>Sheet2!$B$1`; operand `Sheet2!$A$1`; `'My Sheet'!A1`; a reference to a deleted sheet; `dynamicRefresh` after editing Sheet2 without saving. | `formula:cross-sheet-reference`, `-quoted-name`, `-missing`, `formula:cellis-operand-cross-sheet` |
| `stage5-cfvo-formulas` | Scale `formula` `=$F$1` and `=MIN($A$2:$A$9)+1`; `num` `$F$1`; `percent`/`percentile` `$F$1`; bar `formula` min/max; icon `formula` and `num` `$F$1`; relative `=A2` cfvo; cfvo results text/blank/error (hypotheses: withheld / 0 / skipped). | `scale:thresholds-formula`, `bar:bounds-formula`, `icon:thresholds-formula`, `formula:cfvo-reference`, `-expression`, `-relative`, `-nonnumeric` |
| `stage5-cfvo-scope` | `sqref="B2:B5 D2:D5 F2:F5"` scale and icon set with a relative `formula` bound `=MAX(B2:B5)`, plus an absolute-bound twin (`=$H$1`) as a negative control. The first range B2:B5 is the **control**: both hypotheses evaluate it at the same anchor. D2:D5 and F2:F5 hold maxima different from B's, so rule-wide (one value from B2's anchor) and per-range (each range's own anchor) differ there; each of D and F has ≥ 2 decisive cells (round-2 #7). | `formula:cfvo-scope-multi-range` |
| `stage5-x14-scale` | The dialog-authored cross-sheet scale: does Excel write `<x14:colorScale>` or the base form? Capture if x14. | `scale:dialect-standalone` |
| `stage5-active-present` | One workbook per visual kind: hand-authored x14 `activePresent="1"` plus a rule-level `<xm:f>` false on ≥ 2 cells; record retain/strip/repair. | `formula:active-present-scale`, `-bar`, `-icon` |
| `stage5-unresolved-shapes` | One workbook each: `=MyName>0`, `=[1]Sheet1!A1>0`, `=Table1[@Col]>0`, `=INDIRECT("A1")>0`, `={1,2}`, `=(A1,B1)`, `=A1#`, a formula on `duplicateValues`. Records Excel's outcome; lifts nothing. | — |

- [ ] **Conformance first (Lesson 1).** Add `stage 5 formula cells are withheld until their evidence lands` to `tests/xlsx-cf-oracle-conformance.test.ts`. Under production evidence, every Stage 5 cell's `CfPrepared.at(...)` is either not `ready` or carries no formula-sourced property. FINDINGS stay withheld-only (extending the guard at `:1395`). The Stage 1–4 assertions stay byte-unchanged except those re-pointed by S5-STUB in Task 35.
- [ ] Write `stage 5 cases join the manifest and stage 1–4 cases are unchanged`. Manifest ids equal the combined `CF_CASES`, and the Stage 1–4 capability ids and sets are unchanged. RED, then GREEN.
- [ ] Write `formula marker pins rule formula, operands, cfvo f, dialect and activePresent exactly`. The verifier refuses each mutation: `$B2` → `B2`; `>` → `>=`; operand order swapped; `cfvo.f` `$F$1` → `$F$2`; x14 scale moved to base; `activePresent` dropped; `sqref` order changed when the marker pins Excel's resaved order; a formula moved to another part. It compares semantics: Excel's canonical reference spelling and sheet quoting, not raw strings.
- [ ] Write `semantic DisplayFormat evidence needs two decisive cells and a control (Lesson 3)`. On synthetic sidecars:
  - a configuration whose hypotheses agree at every designated cell → `unverified-nondiscriminating`;
  - one decisive cell → `replicationShortfall` refusal;
  - two agreeing decisive cells with one hypothesis matching → verified;
  - a positive whose control is also formatted → refused.
- [ ] Write `audit --stage 5 --write is idempotent (Lesson 4)`.
  - Two runs give a byte-identical `verifiedCapabilities.ts`.
  - Without `--publish-stage5` the Stage 5 export is carried over verbatim (absent at first).
  - `--publish-stage5` writes `VERIFIED_STAGE5_CAPABILITIES` only when `stageComplete=true`.
  - Stage 1–4 exports are unchanged.
- [ ] Write `a publication requires production's reading to equal Excel's matched reading`. A capability whose matched reading (for example `range-top-left`) is not the reading production implements stays withheld, as Stage 4's `CF_STAGE4_PRODUCTION_READINGS` does.
- [ ] Write `the manual-calculation route emits no premature calculate (round-1 #4)`. For a `calcObservation: 'pre-calculate'` case, the generated AppleScript of build and capture contains no `calculate` before the first observation line, and records the application calculation mode; the explicit-calculation observation is a separate, later script. A mutant that keeps `build.ts:180` fails.
- [ ] Extend the osacompile test with the Stage 5 probe scripts under `it.skipIf(!hasOsacompile)`. The clock cases reuse the TZ pin (Lesson 8).
- [ ] **Orchestrator:**
  1. Author `stage5-formula-authored` through the dialog, save, and diff its XML into the templates (Lesson 12).
  2. Run `bun scripts/cf-refs/build.ts --stage 5 --case <id>` and `capture.ts --stage 5 --case <id>` per case, one malformed shape per workbook. A `calcObservation: 'pre-calculate'` case skips the unconditional `calculate` in the build resave (`scripts/cf-refs/build.ts:180`) and the clock probe's `calculate` (`capture.ts:66`), and runs `buildManualObservationScript` first.
  3. Run the clock cases with `--refresh-clock-date`.
  4. Record every strip or repair as the case outcome and never retry it.
  5. Close owned workbooks between cases. Never manufacture evidence.
- [ ] Inventory the corpus for formula-bearing CF (ZIP XML; skip `~$`) into `docs/validation/xlsx-cf.md` §4. If the count is 0, as §4 recorded for all CF in `corpus/`, record that, and make the Excel-authored fixture the corpus gate (Task 39).
- [ ] Run `npx vitest run tests/xlsx-cf-oracle.test.ts tests/xlsx-cf-oracle-conformance.test.ts`, `npx tsc --noEmit -p tsconfig.scripts.json`, `npx tsc --noEmit -p .` and `git diff --check`. Commit (`test(xlsx-cf): add Stage 5 formula oracle cases`), report, review.

### Task 32: Formula compile, admission, allow-list, model and both-path parsing (Claude Sonnet; review per table)

**Files:** Create `src/xlsx/cf/formula.ts` (compile and admission half, and the evaluation-context contract), `tests/xlsx-cf-formula-parse.test.ts`, `tests/xlsx-cf-evaluation-context.test.ts`. Modify `src/xlsx/cf/model.ts`, `parse.ts`, `evidence.ts`, `service.ts` (`statistical`, `estimateCfPrepareBytes` arms only), `render.ts` and `text.ts` (the `evaluation` option), `src/xlsx/view/pagedSource.ts` (`CfRemotePrepareOptions.evaluation`), `src/xlsx/storage/pagedImport.ts` (`calcMode`), `src/worker/viewerSession.ts` (metadata), `src/render/paint.ts:425`, `src/components/xlsx/sheetAssets.ts:259`, `InteractiveSheetDoc.tsx` (pass-through only), `src/xlsx/cfLoader.ts` (`canRuleMatchBlanks`), `tests/xlsx-cf-oracle-conformance.test.ts` (callers at `:213`), `tests/xlsx-cf-paged-parse.test.ts`, `tests/xlsx-cf-lazy.test.ts`.

**Interfaces:**
- **Consumes:** `parseFormula`, `tokenize`, `AstNode`, `FUNCTIONS`, `readFormulas`, `cfvoLiteral`, `anchorOf`, `CfRuleOrigin.formulas`, `CF_LIMITS.formulaTokens`, `hasEvidence`.
- **Produces:** the Task 32 block of **Produces**, including the evaluation-context contract and its wiring through every route **before** Task 35 evaluates any formula (round-2 #3). The `expression` stub at `parse.ts:1526` and the operand stub at `:1449–1451` are replaced; cfvo stubs stay until Task 37.

- [ ] Write `the evaluation context is built once per document and the clock once per preparation (round-2 #3)`. `cfDocumentEvaluationOf` called twice on one document returns equal policies; two preparations take two clock snapshots from the injected `now`. An undefined `doc.valuePolicy` resolves `'engine'` and an absent `calc` resolves `'auto'`. Every route (`prepareCfPaint`, `prepareCfTextScope`, `prepareCfRemoteOnce`, `getPaintables`, `prepareSheetAssets`, the interactive text and paint scopes, the worker metadata) forwards the policy; a route given none leaves formula rules `cf-evaluation-context-missing` and Stage 1–4 rules unchanged. The conformance callers (`tests/xlsx-cf-oracle-conformance.test.ts:213`) pass `cfDocumentEvaluationOf(doc)`, and their Stage 1–4 assertions stay green. A source-text assertion confirms that the eager sites import no CF value and reach the builder only through the loaded lazy module.

- [ ] Write `compile admits the allow-list and refuses every out-of-scope node statically (S5-B)`. Each of these → `{ refused: 'cf-unsupported-rule' }`: `MyName>0`, `[1]Sheet1!A1>0`, `Table1[@Col]>0`, `INDIRECT("A1")`, `IF(FALSE,INDIRECT("A1"),1)` (the untaken branch), `{1,2}`, `(A1,B1)`, `A1#`, `Sheet1:Sheet3!A1`, `NOW()`, `VLOOKUP(A1,B:C,2)`, `SEARCH("a",A1)`. `LEN(CONCAT(A1:A64))`, `CONCAT(A1:A3)`, `A2:A4>0`, `IF(TRUE,A2:A4,0)`, `IFERROR(A2:A4,0)` and `NOT(A2:A4)` are refused too (round-2 #1, #4); for `LEN(CONCAT(A1:A64))` a spy proves that no snapshot read and no `textChars` call happens. `MOD(A1,2)=1`, `COUNTIF($A$1:$A$50,A1)>1`, `CONCAT(A1,"x")` and `AND(A1>0,B1<5)` compile. After `registerFormulaFunction('ZZZ', …)`, `ZZZ(A1)` still refuses.
- [ ] Write `functions with implicit expansion compile only at literal geometry (S5-SUMIF-SHAPE)`. `SUMIF(A1:A2,"",B1)`, `AVERAGEIF(A1:A2,"",B1)` and `SUMIF(A:A,"",B1)` are refused at compile; `SUMIF(A1:A2,"",B1:B2)`, `AVERAGEIF(A1:A2,"",B1:B2)`, `SUMIF(A:A,"",B:B)` and `SUMIF(A1:A2,"")` compile, and their footprint holds the effective geometry (B1:B2, B:B). `SUMIFS`, `AVERAGEIFS` and `COUNTIFS` compile with literal geometry. Every `CF_FORMULA_FUNCTIONS` entry declares `geometry` for every argument (table test against the sweep).
- [ ] Write `compiled formula records atoms, families, invariance and clock use`. `=$A$1>AVERAGE($A$2:$A$9)` → `invariant: true`, atoms `[A1 abs, A2:A9 abs]`, families `['aggregate']`. `=A2>0` → `invariant: false`. `=ROW()=2` → `invariant: false`. `=A2=TODAY()` → `usesClock: true`. `=Sheet2!$B$1<A2` → `dependencySheets: ['Sheet2']`.
- [ ] Write `compile caps refuse before construction`: 4,097 nodes → `cf-resource-limit`; 65 atoms → `cf-resource-limit`; 17 referenced sheets → `cf-resource-limit`. Each limit itself is admitted.
- [ ] Write `compile reserves before parsing and releases with the rule (round-1 #3)`. At `tokens × CF_FORMULA_NODE_BYTES − 1` free bytes, `parseFormula` is never called (spy) and no AST exists; at exactly that, it compiles and the parse account holds the reservation. A parse failure, a cancelled streamed import and sheet disposal each return the parse ledger to baseline. A test with the accounting removed (mutant: no reserve) fails the first assertion.
- [ ] Write `expression and operand rules parse to formula definitions behind evidence`. Under the **complete** injected parse set (`formula:dialect-base`, `formula:anchor-single`, `formula:anchor-mixed-abs`; round-2 #6), `<cfRule type="expression"><formula>$B2>$C2</formula>` over B2:D9 yields `{ kind: 'expression', sourceType: 'expression' }` with anchor B2. Without evidence → `cf-unsupported-rule`, writes = the DXF writes, one diagnostic. `cellIs greaterThan <formula>$F$1</formula>` yields an operand `{ formula }`, and literal `100` still yields a `CfLiteral` (Stage 1 unchanged). A divergent `containsBlanks ISBLANK(D2)` yields an `expression` definition with `sourceType: 'containsBlanks'`. The canonical `LEN(TRIM(D2))=0` keeps its predicate definition.
- [ ] Write `formulaRequiredCapabilities splits parse and evaluation gates (round-1 #12)`. With rule context `{ dialect: 'base', rangeCount: 1, sourceType: 'expression' }`, `=$B2>$C2` gives `parse = {formula:dialect-base, formula:anchor-single, formula:anchor-mixed-abs}` and `evaluation` ⊇ the family ids plus the truth reading of its root shape. `rangeCount: 2` adds `formula:anchor-multi-range` to `parse`; `dialect: 'x14'` swaps in `formula:dialect-x14`; a cross-sheet atom adds `formula:cross-sheet-reference`; `TODAY` adds `formula:fn-today`; `use: 'cfvo'` with `rangeCount: 2` adds `formula:cfvo-scope-multi-range`. Neither set is ever empty.
- [ ] Write `resident and streamed formula rules are deep equal at every chunk boundary`, extending `tests/xlsx-cf-paged-parse.test.ts:32` with an expression, a formula operand, a divergent predicate and a quoted cross-sheet name, with boundaries inside `<formula>` and `<xm:f>`.
- [ ] Write `formula definitions are priced and statistical-aware`. `estimateCfPrepareBytes` adds `CF_FORMULA_NODE_BYTES × nodes` per formula rule (source-text assertion: no new integer literal on the formula arm). `statistical()` is unchanged for expressions (they use the snapshot, not statistics).
- [ ] Write `blank synthesis stays on for formula rules (S5-P)`. `canRuleMatchBlanks` returns true for resolved `expression`, formula-operand `cellIs` and divergent predicates, and stays true for the unresolved `expression`. Record the eager gzip delta.
- [ ] Add `'formula'` and `'dependencies'` to the file list of `excludes every CF implementation from both ordinary static closures` (`tests/xlsx-cf-lazy.test.ts:14–17`). Assert from the module map that `indexModuleSetHash` (`scripts/eager-baseline.json`) is unchanged by importing the engine from the `cf` chunk.
- [ ] Run `npx vitest run tests/xlsx-cf-formula-parse.test.ts tests/xlsx-cf-parse.test.ts tests/xlsx-cf-paged-parse.test.ts tests/xlsx-cf-scale-parse.test.ts tests/xlsx-cf-bar-parse.test.ts tests/xlsx-cf-icon-parse.test.ts tests/xlsx-cf-lazy.test.ts` (RED, implement, GREEN), tsc and the eager gates. Commit (`feat(xlsx-cf): compile and admit CF formulas`), report, review.

### Task 33: Canonical and divergent `timePeriod` formula classification (Claude Sonnet; review per table)

**Dependencies:** Task 32 (compile path) and the Task 31 `stage5-predicate-formulas` case. Runs before Task 35 so the divergent arm exists when expressions resolve.

**Files:** Modify `src/xlsx/cf/parse.ts` (the `timePeriod` arm at `:1484–1492`), `tests/xlsx-cf-parse.test.ts` (R1c block `:257–262`), `tests/xlsx-cf-paint-fixes.test.ts` (`:165`), `tests/xlsx-cf-paged-parse.test.ts`. Create `tests/xlsx-cf-timeperiod-formula.test.ts`.

**Interfaces:**
- **Consumes:** `refsMatchAnchor`, `stripOutsideStrings` (`parse.ts:341,380`), `PERIODS` (`:124`), the Stage 1 clock workbooks' resaved XML, `compileCfFormula` (Task 32).
- **Produces:** `canonicalTimePeriodFormula` (new); the divergent `timePeriod` arm.

- [ ] Write `canonical timePeriod patterns come from Excel's own XML`. For each of the ten periods, the formula Excel wrote in the resaved `stage1-timeperiod-today-clock` and `-2` workbooks is read by the test from the saved archive and must classify canonical; the pattern table is not typed by hand in the test.
- [ ] Write `divergent timePeriod formulas never reach the attribute predicate (round-1 #9)`. `timePeriod="today"` with `<formula>A2=TODAY()-3</formula>` or `FLOOR(B2,1)=TODAY()-2` is unresolved (`cf-unsupported-rule`) on the resident and streamed paths, writes its DXF mask, and has one diagnostic. Under the complete injected set for the "formula" reading (`formula:predicate-timeperiod-divergent`, `formula:dialect-base`, `formula:anchor-single`, `formula:fn-today`, `formula:truth-boolean`), `A2=TODAY()-3` compiles as `{ kind: 'expression', sourceType: 'timePeriod' }`, and `FLOOR(B2,1)=TODAY()-2` stays `cf-unsupported-rule` because `FLOOR` is neither registered nor allow-listed (round-2 #6). Under the "attribute" reading the divergent formula keeps the attribute predicate. An absent formula and a foreign-namespace formula keep their main-branch outcomes.
- [ ] Rewrite the R1c tests (`tests/xlsx-cf-parse.test.ts:257–262`, `tests/xlsx-cf-paint-fixes.test.ts:165`) to the new expectation, citing S5-G. Also update the ten-period table at `tests/xlsx-cf-parse.test.ts:231–241`: its `lastWeek` row (`:240`) holds the off-by-one `>(WEEKDAY(TODAY())-1)`/`<(WEEKDAY(TODAY())+6)` form, which now classifies divergent; it is replaced by the form Excel wrote in the saved clock fixture (`>=(WEEKDAY(TODAY()))`/`<(WEEKDAY(TODAY())+7)`, as in `tests/xlsx-cf-paint-fixes.test.ts:152`), read from the fixture rather than typed. Every Stage 1 clock conformance cell stays byte-unchanged (`tests/xlsx-cf-oracle-conformance.test.ts:1411` passes unmodified).
- [ ] Extend `resident and streamed CF are deep equal at every chunk boundary` with canonical and divergent `timePeriod` formulas.
- [ ] Run `npx vitest run tests/xlsx-cf-timeperiod-formula.test.ts tests/xlsx-cf-parse.test.ts tests/xlsx-cf-paint-fixes.test.ts tests/xlsx-cf-paged-parse.test.ts tests/xlsx-cf-oracle-conformance.test.ts` (RED, implement, GREEN), tsc and the eager gates. Commit (`fix(xlsx-cf): classify timePeriod formulas instead of trusting the attribute`), report, review.

### Task 34: Translator, truth readings, same-sheet dependency snapshot and budgets (Claude Opus; review per table)

**Files:** Complete `src/xlsx/cf/formula.ts` (including `planCfHoist`, `cfCountifKey`); modify `src/xlsx/cf/statistics.ts` (`prepareHoists`) and `service.ts` (the statistics loop runs hoists beside statistical rules); create `src/xlsx/cf/dependencies.ts`, `tests/xlsx-cf-formula-hoist.test.ts`, `tests/xlsx-cf-formula-eval.test.ts`, `tests/xlsx-cf-formula-budget.test.ts`. Modify `src/xlsx/cf/limits.ts`, `tests/xlsx-cf-active-units.test.ts`, `scripts/cf-refs/active-units.ts`.

**Interfaces:**
- **Consumes:** `evaluateNode` (never `evaluateFormulaInternal`, S5-D), `createReferenceServices`, `SparseCellStore`, `MaterializationBudget`, `ResolvedSemantics`, `CfValueSource.pages`/`bounds`, `createCfWorkBudget`, `createCfMemory`, `measurePayloadBytes`.
- **Produces:** the Task 34 block of **Produces**. Evaluation stays pure: no DOM, no wall clock (the clock arrives through `semantics`).

- [ ] Write `translation follows the measured edge mode and pinned axes`. `=A2>0` anchored at B2, evaluated at B5 → reads A5. `$A2` at C5 → `$A5`. `A$2` → row 2. At row 1048576, `A1048577` → A1 under `'wrap'`, `#REF!` under `'ref-error'`, and unresolved with no edge evidence (S5-C). The translated tree is reserved before allocation: at `nodes × CF_FORMULA_NODE_BYTES − 1` nothing is built.
- [ ] Write `footprint over a target rect is the shifted union, clamped`. `=A2>B1` anchored at C2, over C2:C9 → `{A2:A9, B1:B8}`. `COUNTIF($A:$A,A2)` on a 50-row sheet → `A1:A50` plus `A2:A9`. 262,145 cells → `'over-cap'`.
- [ ] Write `shape is gated before projection (S5-D, round-1 #7)`. `=A2` over an unstored A2 yields a `blank` result, never `0`, and with only `formula:truth-number` injected it is unresolved. `=A1:A3` at row 2 yields a `range` result; with only number evidence it is unresolved, and under `range-implicit` vs `range-top-left` it reads A2 vs A1. `evaluateFormulaInternal` is never imported by `formula.ts` (source-text assertion). Under each injected reading the observable outcome matches S5-D; `coerceToBoolean` is never called.
- [ ] Write `values a formula reads equal the host literals (S5-E)`. For blank, `""`, `#DIV/0!`, a date serial, TRUE and `"1"`, `cfLiteralToElement` and `cfElementToLiteral` round-trip. `=ISBLANK(A2)` matches the unstored blank and not `""` (Excel-backed: `cases.ts:981–982`). `=A2=""` is asserted only against the `stage5-formula-functions` observation, never against the engine's own answer (round-1 #12).
- [ ] Write `a recorded unsupported feature discards a masked result (round-1 #1)`. p1 `IFERROR(COUNTIF(A:A,TRUE),TRUE)` with a red fill and p2 `cellIs > 0` with a green fill and a bold font over the same cells: the p1 result is unresolved (`cf-formula-unsupported`) although the engine returns `true`; the cell shows neither red nor green, and bold only if the p1 DXF does not write font. The same holds for a cycle flag and for a materialization refusal swallowed by `IFERROR`.
- [ ] Write `a read outside the snapshot taints and discards (Review Focus 1)`. A formula whose dynamic branch reads a cell outside the footprint returns unresolved `cf-formula-dependency`, even when its computed value would be `true`. A cached-unavailable cell under `'cached'` → unresolved. No path returns a blank for a missing cell.
- [ ] Write `relative-range traversal work is charged per occurrence before it happens (S5-N, round-1 #2, round-3 #6, round-4 #2)`. Not hoisted (relative axes). Each expectation is derived from the component breakdown `cfCallWorkUnits` reports, and the test asserts each component:
  - `SUM(A1:A4)+SUM(A1:A4)` on a 4-cell view: 5 nodes + 4 view build + 2 × (4 enumeration + 4 reads + 8 sort) = 41 ≤ 128. With the budget at 40 the call never runs (spy) and the rule is unresolved, while a second rule in the batch still evaluates.
  - `COUNTIF(A$1:A$4,B2)` (column relative, so not hoisted) over the 1,000 numeric targets B2:B1001. The snapshot is explicitly bounded: A1:A4 plus B2:B1001, 1,004 cells, built once for 1,004 units. Each call's restricted view holds 5 cells (A1:A4 and its own criterion cell), so its bound is 3 nodes + 5 view build + 5 enumeration + 4 range reads + 1 criterion read + 15 sort (5 · ⌈log₂ 5⌉) + 4 numeric comparisons = 37, not the 1,004-cell enumeration of the shared snapshot. With an injected preparation ceiling of 5,000 units, the snapshot build leaves 3,996, so exactly ⌊3,996 / 37⌋ = 108 targets (B2:B109) evaluate and B110:B1001 are unresolved in row-major order.
  - With charging removed (mutant), both tests fail.
- [ ] Write `SUMIF/AVERAGEIF read the same cells on the view as on the full store (S5-SUMIF-SHAPE regressions)`. With A1:A2 blank, B1 = 0 and B2 = 7: `SUMIF(A1:A2,"",B1)=0`, `AVERAGEIF(A1:A2,"",B1)=0` and `SUMIF(A:A,"",B1)=0` are each refused (compile), never evaluated on a view. With matching-size sum ranges (`B1:B2`, and `B:B` for the whole column), each evaluates on the per-call view to exactly the engine's result on the full store (`SUMIF(A1:A2,"",B1:B2)` = 7, so `=0` is false). A `viewFor` request whose effective area is not fully covered by the snapshot is refused before any `stored()` visit (spy).
- [ ] Write `full-store enumeration is counted per traversal occurrence (round-4 #1)`. On a 6-cell view, `SUM(A1:A6)+SUM(A1:A6)+SUM(A1:A6)+SUM(A1:A6)` costs 11 nodes + 24 matched reads + 24 enumeration visits + 72 sort units = 131, plus 6 to build the view = 137 > 128, so it is refused before invocation (spy: no `evaluateNode` call, no `stored()` visit). A mutant that omits the enumeration term computes 113 and admits it, and fails the test.
- [ ] Write `each synchronous call stays within one checkpoint and the loop yields between cells (round-2 #2)`. A formula whose per-call bound is 128 evaluates; at 129 it is refused at footprint time and `evaluateNode` is never called (spy). Over 1,000 targets, `hostTask` is awaited between cells whenever the budget says `'yield'`, never inside a call (spy on call depth). An abort scheduled from a host task fired at the 10th checkpoint stops the preparation before the next cell's call: no further `evaluateNode` call, the preparation throws `cancelled`, and the ledger returns to baseline.
- [ ] Write `absolute-range aggregates are hoisted and paint on a 10,000-row column (S5-HOIST)`. On a populated 10,000-row column A, under complete injected evidence, `COUNTIF($A:$A,A2)>1` and `A2>AVERAGE($A:$A)` paint exactly the cells the engine's own per-cell evaluation of the same formula marks (computed once offline on the snapshot as the expectation). Every per-cell call is ≤ 128 units (spy), the hoist scan awaits `hostTask` at its checkpoints and pays from a scan budget, not the per-cell ceiling (spy), the snapshot holds only the per-cell footprint (A2 and its neighbours, not the column), and the column is scanned once per rule (source-visit counter). A hoist-free mutant refuses both rules (`cf-resource-limit`).
- [ ] Write `the COUNTIF frequency map equals the engine's COUNTIF on admitted inputs (S5-HOIST property test, round-3 #5)`. Random **admitted** snapshots (numbers, ASCII text in mixed case, blanks, errors) and random admitted criterion cells: `map.get(cfCountifKey(criterion)) ?? 0` equals the engine's `COUNTIF(range, criterion)` on the same snapshot, with at least 500 successful comparisons each for number criteria and for text criteria (the test asserts the counts), plus fixed deterministic cases: an empty range, a range of only errors, duplicates differing only in case. A mutant map that admits wildcard criteria (`"a*"`) fails.
- [ ] Write `COUNTIF hoist policy refusals (S5-HOIST, round-3 #1, #5)`. Deterministic, and separate from the engine-equivalence test: a `"#N/A"` criterion over a range holding `#N/A` is refused per cell (the engine counts 1, a string lookup would count 0); a 256-character ASCII criterion is refused (`conditional-long-criterion-unverified`); a 123-character ASCII criterion against a range holding one 32,767-character string is refused (`conditional-wildcard-resource-limit`); `" 1 "` and `"1"` as criteria are refused; ranges with `" 1 "`, a boolean or non-ASCII text refuse the whole rule. Each asserts refusal without consulting the engine's result.
- [ ] Write `hoisted scalar aggregates equal the engine's own evaluation on admitted inputs (S5-HOIST property test, round-3 #3, #5)`. Random snapshots over mixed cell kinds (numbers, numeric text, text, booleans, blanks, errors), including empty and all-error ranges: for each hoisted `SUM`, `COUNT`, `COUNTA`, `AVERAGE`, `MIN`, `MAX`, and invariant `COUNTIF`/`SUMIF` with literal and absolute-reference criteria, every input the hoist admits gives exactly the engine's value or error code, with at least 300 successful comparisons per profile (asserted). Fixed deterministic cases: `SUM` of `[1e16, 1, -1e16]` = 0 (the compensated statistics sum gives 1, so a mutant reusing it fails); an overflowing `SUM` gives the engine's overflow result; the first error in row-major order wins over a later one; an empty range gives the engine's empty result for each function. Inputs the hoist refuses are tested separately by policy. A mutant that ignores booleans, and one that ignores errors, each fail.
- [ ] Write `a cancelled or refused hoist leaves nothing charged (S5-HOIST)`. An abort mid-scan, and a frequency-map growth one byte over its reservation (old plus new table at the peak), each publish nothing, return the ledger to baseline and leave the rule unresolved. A host-sheet commit, and a referenced-sheet commit (`dependencyGen`), each invalidate the cached hoist and every tile of the sheet.
- [ ] Write `text work is bounded by measured payload before invocation (round-2 #1, round-3 #6)`. Two 32,767-character strings make `LEN(A1&A2)` refused at compile-from-payload time **whatever memory is available** (its text units exceed the 128-unit cap; spy: no `evaluateNode` or `textChars` call, even with an unlimited scratch grant). A genuinely sub-cap positive, `LEN(A1&A2)` over two 10-character strings, runs at its exact scratch bound and never runs one byte below it. The matcher bound: `COUNTIF(A1:A2,B1)` over 10-character strings has a bound of 2 × 11 × 11 = 242 transitions and is refused; `COUNTIF(A1:A2,B1)` over 3-character strings runs: 3 nodes + 3 view build + 3 enumeration + 3 reads + 6 sort + 32 matcher transitions (2 × 4 × 4) = 50 ≤ 128. The sub-cap `LEN(A1&A2)` over 10-character strings totals 4 nodes + 2 view build + 2 reads + 20 text units = 28 ≤ 128, with enumeration counted (it has no range traversal). A mutant that drops the text or matcher term fails.
- [ ] Write `evaluator scratch and snapshot growth are reserved before use (Lessons 9, 15, round-1 #3)`. With free bytes one below `cfEvaluatorScratchBytes(f, metrics)`, where `metrics` is the `CfPayloadMetrics` returned by `viewFor` for that call, `evaluateNode` is never called (spy) and nothing is constructed. A snapshot growth from 1,024 to 2,048 entries charges old plus new at the peak. A refused shared snapshot retried per rule uses a fresh per-attempt budget (spy: two distinct budgets) and both draw from one preparation ceiling. Cancellation between calls, or during the asynchronous snapshot build or hoist scan, and dispose each return the ledger to baseline.
- [ ] Write `formula results stay finite near ±1e308 (Lesson 6)`. `=A2*10>0` with A2 = 1e308 → `#NUM!` in the engine → unresolved, never `Infinity` as true.
- [ ] Extend `tests/xlsx-cf-active-units.test.ts`. The largest production call with formula rules (1,000 text-path and 256 paint-path cells, four formula rules each, one `COUNTIF($A:$A,…)`) spends ≤ `CF_LIMITS.activeUnits / 5`. Re-run `npx tsx scripts/cf-refs/active-units.ts`, and record the numbers in `docs/validation/xlsx-cf.md` §11.
- [ ] Run `npx vitest run tests/xlsx-cf-formula-eval.test.ts tests/xlsx-cf-formula-budget.test.ts tests/xlsx-cf-active-units.test.ts tests/xlsx-cf-budget.test.ts` (RED, implement, GREEN), tsc and the eager gates. Commit (`feat(xlsx-cf): evaluate CF formulas over a bounded snapshot`), report, review.

### Task 35: Formula DXF rules: expression, operands, divergent predicates, composition, stub successors (Claude Opus; review per table)

**Files:** Modify `src/xlsx/cf/service.ts` (formula arm in `evaluate`, snapshot before the page loop), `predicates.ts` (operand path into `evaluateCellIs`), `compose.ts` (only if a gate is missing), `tests/xlsx-cf-oracle-conformance.test.ts` (S5-STUB re-pointing, one commit). Create `tests/xlsx-cf-formula-rules.test.ts`, `tests/xlsx-cf-formula-compose.test.ts`.

**Interfaces:**
- **Consumes:** Tasks 32–34; `evaluatePredicate`, `evaluateCellIs`, `composeCellAsync`, `visualSlotOf`, `CF_STAGE5_REUSED_CAPTURES`.
- **Produces:** resolved formula outcomes in `CfCellState.properties` on both adapters; the re-pointed stub assertions.

- [ ] Write `expression rule paints exactly the DisplayFormat cells (Lesson 1)`. For every verified Stage 5 case, production's per-cell match equals the sidecar's `excel-display-format` observation. Every pending, refused or non-discriminating cell asserts no formula-sourced property. No expectation is computed by the function under test.
- [ ] Write `formula operands reuse the Stage 1 comparators`. `cellIs > $F$1` with F1 = 100 separates 99/100/101 exactly as literal 100 does under the same `cellIs:*` capabilities. An operand `#N/A` → unresolved on that cell only; a reversed `between` follows `cellIs:between-reversed`.
- [ ] Write `divergent predicates follow the formula (S5-G)`. `ISBLANK(D2)` matches the unstored blank but not `""` (Stage 1 cells `cases.ts:981–982`). `NOT(ISBLANK(G2))` matches 0 (`:988`). `ISNA(J2)` matches `#N/A` and not `#DIV/0!` (`:994–995`). The `SEARCH("zz",J2)` cells stay withheld (S5-B).
- [ ] Write `formula composition across kinds (S5-O)`:
  - an expression p1 `stopIfTrue` true halts a lower DXF, scale, bar and icon only with `formula:composition-stop-ahead-<kind>`, else the lower contribution is withheld;
  - an expression fill vs a scale fill follows `formula:composition-expression-scale`;
  - two overlapping expressions write first-writer-wins only with `formula:composition-overlapping-expressions` (and the Stage 1 multi-positive guard);
  - an unresolved expression p1 taints exactly its DXF writes;
  - all Stage 2–4 compose suites (`tests/xlsx-cf-scale-compose.test.ts`, `xlsx-cf-bar-compose.test.ts`, `xlsx-cf-icon-compose.test.ts`, `xlsx-cf-bot-fixes.test.ts`) stay green unchanged.
- [ ] Write `stub successors repaint only where Excel did (S5-STUB)`. For every row of the S5-STUB inventory, each conformance cell is Excel's recorded outcome under its named successor capabilities, or withheld; `stage1-blanks-errors` D2/G6/J8/M6 match their recorded Excel fills under `formula:predicate-divergent`. The unit-test rows assert the same refusal under `withExactCfEvidence([])`. A grep test asserts the re-pointing commit touches only the inventoried lines. `git diff main -- src/xlsx/cf/verifiedCapabilities.ts` is empty at this task.
- [ ] Write `formula rules evaluate identically on the resident and paged adapters`. On the Stage 5 fixtures, `createResidentCfSource` and the paged `readConditionalFormatting` produce deep-equal per-cell states.
- [ ] Run `npx vitest run tests/xlsx-cf-formula-rules.test.ts tests/xlsx-cf-formula-compose.test.ts tests/xlsx-cf-oracle-conformance.test.ts tests/xlsx-cf-compose.test.ts tests/xlsx-cf-predicates.test.ts tests/xlsx-cf-worker.test.ts` (RED, implement, GREEN), tsc and the eager gates. Commit (`feat(xlsx-cf): evaluate expression rules and formula operands`), report, review.

### Task 36: Cross-sheet values, pending pages, `dependencyGen` and remote invalidation (Claude Opus; review per table)

**Files:** Modify `src/xlsx/cf/dependencies.ts` (`CfWorkbookSource` adapters), `model.ts`, `cache.ts`, `render.ts`, `text.ts`, `src/xlsx/view/contracts.ts`, `pagedSource.ts`, `src/worker/viewerRemoteCf.ts`, `viewerRemoteSource.ts`, `viewerSession.ts`, `src/components/xlsx/InteractiveSheetDoc.tsx`. Create `tests/xlsx-cf-formula-cross-sheet.test.ts`; extend `tests/xlsx-cf-live.test.tsx`, `tests/xlsx-cf-worker.test.ts`.

**Interfaces:**
- **Consumes:** `CfStamp`, `sameCfStamp`, `CfReadResult`, `RemoteCfCache`, `workbook.store.pages` and `importedThroughRow` (the `fillCharts` pattern, `pagedSource.ts:1198,1231`), `WorkbookOwner.sheetVector` (`contracts.ts:777`), `cfLiveKey` (`InteractiveSheetDoc.tsx:1083`).
- **Produces:** the Task 36 block of **Produces**. The construction sites at `InteractiveSheetDoc.tsx:1168,1434`, `render.ts:313–320`, `text.ts:428–436`, `pagedSource.ts:982–990` and `viewerSession.ts:3329,3590,3604,3627` set `dependencyGen` (0 when a sheet has no cross-sheet rule).

- [ ] Write `a referenced-sheet commit makes the old batch stale synchronously (Review Focus 4)`. Sheet1 `=Sheet2!A2>0`, painting green at g1. Commit Sheet2 A2 = −1: in the same tick, `at()` returns `cf-stale-generation` on resident, store and worker legs; no tile paints green after the commit; the g2 preparation paints base. A late g1 completion never publishes.
- [ ] Write `cross-sheet values and generation come from one overlay-aware store (round-1 #6)`. Sheet2 A2 saved `5`, pivot overlay `-5`. Sheet1 `=Sheet2!A2>0`: before the commit the rule matches; after the overlay commit it does not; after the pivot release it matches again. At each transition the snapshot value and the stamp come from the same `PagedXlsxSheetAdapter` (worker) and `source.store` (resident). A commit landing between snapshot build and adoption cancels the preparation (`cf-stale-generation`); a never-activated Sheet2 is created through `storeFor` and reused on activation (one adapter).
- [ ] Write `dependencyGen strictly increases on any referenced commit and ignores unrelated sheets`. With Sheet2 and Sheet3 referenced, a Sheet3 commit changes it and a Sheet4 commit does not. A host commit still invalidates through `workbookValueGen`. Wrapping a store without CF never touches its stamp.
- [ ] Write `unimported referenced pages are pending, then repaint once`. A Sheet2 still importing → the Sheet1 cells reading it are `cf-pending-data` with `pendingSheetIds: ['Sheet2']`, no permanent diagnostic, and exactly one invalidation of Sheet1 when Sheet2 completes. A Sheet2 that ends `incomplete` → `cf-read-failed` once.
- [ ] Write `remote cache entries follow their reply's dependencies`. A reply listing `{Sheet2, valueGen 3}` is current until Sheet2's identity moves to 4, at which point `get()` is stale and `viewerRemoteSource` invalidates every tile of Sheet1 (Lesson 5). A pivot commit on Sheet2 (`:706`) and a pivot release (`:752`) both reach it.
- [ ] Write `cross-sheet edge cases refuse identically`. An unknown sheet → `#REF!` only under `formula:cross-sheet-missing`, else unresolved. A quoted `'My Sheet'!A1` resolves. More than 16 referenced sheets → `cf-resource-limit`. A worker that cannot name a referenced sheet's generation → `cf-dependency-generation-unknown`.
- [ ] Write `dependency bytes are measurer-derived`. `CfReadResult.dependencies` at 16 entries is priced by `measurePayloadBytes`. At grant − 1 the formula states are omitted with `cf-grant-required`, base cells return, and `admitReply` never replaces the reply.
- [ ] Write `a no-cross-sheet workbook does zero cross-sheet work`. No `CfWorkbookSource` read, `dependencyGen` stays 0, and the reply carries no `dependencies` key.
- [ ] Run `npx vitest run tests/xlsx-cf-formula-cross-sheet.test.ts tests/xlsx-cf-live.test.tsx tests/xlsx-cf-worker.test.ts tests/xlsx-cf-lifecycle.test.ts tests/worker-package-bootstrap.test.ts` (RED, implement, GREEN), tsc and the eager gates (record the `InteractiveSheetDoc` delta). Commit (`feat(xlsx-cf): evaluate cross-sheet CF formulas`), report, review.

### Task 37: Formula cfvo bounds, x14 colour scales and `activePresent` (Claude Opus; review per table)

**Files:** Modify `src/xlsx/cf/parse.ts` (`parseColorScaleRule`, `barBound`, `iconThreshold`, the activation checks), `colorScale.ts`, `dataBar.ts`, `iconSet.ts`, `service.ts`. Create `tests/xlsx-cf-formula-cfvo.test.ts`; extend the scale/bar/icon parse tests.

**Interfaces:**
- **Consumes:** Tasks 32–36; `resolveScaleThresholds`, `resolveBarThresholds`, `resolveIconThresholds`, `lerp`, `percentileOrderStatistic`, `scaleRequiredCapabilities`/`barRequiredCapabilities`/`iconRequiredCapabilities`.
- **Produces:** formula bounds on all three families; the x14 scale arm when `scale:dialect-standalone` is verified; the activation gate when S5-J's evidence exists, else the recorded permanent refusal.

- [ ] Write `a formula bound resolves once per measured scope (S5-H, round-1 #10)`. A single-range scale with `formula` `=$F$1` (F1 = 50) over {0..100} resolves its min threshold to 50, evaluated once per preparation (spy). `percent` `$F$1` with F1 = 25 resolves to `lerp(min, max, 0.25)`. A bar `formula` max and an icon `formula` threshold resolve the same way. A multi-range rule with any formula bound is unresolved without `formula:cfvo-scope-multi-range`; under each injected reading, memoisation is keyed rule-wide or per range and the disjoint fixture's cells follow it.
- [ ] Write `non-numeric or relative cfvo results withhold`. Text, blank or error → unresolved until `formula:cfvo-nonnumeric`. A relative `=A2` cfvo → unresolved until `formula:cfvo-relative`. A cross-sheet cfvo follows Task 36.
- [ ] Write `num-val references lift only with both capabilities`. `type="num" val="$A$1"` resolves only with `scale:thresholds-num-val-reference` (already verified) plus `formula:cfvo-reference`; with either missing → `cf-unsupported-rule` as on main.
- [ ] Write `x14 colour scales follow the recorded dialect (S5-I)`. With `scale:dialect-standalone` injected, the Excel-authored `<x14:colorScale>` resolves to the same points as its base twin. If Task 31 recorded the base form, the x14 shape stays `cf-unsupported-rule`, and the test pins that refusal with the recorded reason.
- [ ] Write `activePresent follows its recorded outcome (S5-J)`. If Excel stripped it, every `activePresent="1"` visual stays `cf-unsupported-rule` and `formula:active-present-*` sits in `CF_STAGE5_UNSUPPORTED` with the reason. If retained, an activation false at a cell withholds that cell's visual only, and a missing rule-level `<xm:f>` stays `cf-invalid-rule`.
- [ ] Write `formula bounds stay finite and visuals keep their slots`. Bounds at ±1e308 resolve finite or unresolved, never NaN. A formula-bounded bar and icon on one cell follow the Stage 3–4 slot gates unchanged.
- [ ] Run `npx vitest run tests/xlsx-cf-formula-cfvo.test.ts tests/xlsx-cf-scale-parse.test.ts tests/xlsx-cf-bar-parse.test.ts tests/xlsx-cf-icon-parse.test.ts tests/xlsx-cf-scale-eval.test.ts tests/xlsx-cf-bar-eval.test.ts tests/xlsx-cf-icon-eval.test.ts` (RED, implement, GREEN), tsc and the eager gates. Commit (`feat(xlsx-cf): resolve formula cfvo bounds`), report, review.

### Task 38: Calculation policy, clock, every-view agreement, blank synthesis, a11y/copy/find (Claude Sonnet; review per table)

**Files:** Modify `src/xlsx/cf/render.ts`, `text.ts`, `src/xlsx/view/pagedSource.ts` and `src/worker/viewerSession.ts` only where the leg table exposes a gap in the Task 32 wiring. Create `tests/xlsx-cf-formula-views.test.ts`; extend `tests/xlsx-cf-print.test.ts`, `xlsx-cf-export.test.ts`, `xlsx-cf-copy-find.test.ts`, `xlsx-cf-paged-a11y.test.ts`, `xlsx-cf-presentation.test.tsx`, `tests/browser/xlsx-cf.spec.ts`.

**Interfaces:**
- **Consumes:** `prepareCfPaint`, `prepareCfTextScope`, `prepareCfRemoteOnce`, `describeCfWin`, `isValueUnavailable`, `workbookValuePolicy`, the M2 midnight tick (`InteractiveSheetDoc.tsx:566–576`).
- **Produces:** the leg table of S5-P as an executable test, over the context contract and wiring Task 32 already produced (round-2 #3).

- [ ] Write `every route takes one evaluation context or withholds (S5-L, round-1 #5)`. On a workbook with an uncached formula cell in a footprint: `prepareCfPaint`, `prepareCfTextScope`, `prepareCfRemoteOnce`, `getPaintables`, `prepareSheetAssets` and the copy/find and a11y scopes, given the same `CfDocumentEvaluation` and an injected clock, give pairwise-equal outcomes (unresolved under `'cached'`, evaluated under `'engine'`). Each route called without the context leaves every formula rule unresolved (`cf-evaluation-context-missing`) and leaves Stage 1–4 rules unchanged. A paged manual-mode workbook reports `calcMode: 'manual'` from `PagedWorkbook` and withholds exactly as the resident one does. `recomputeDependents` is never called (spy).
- [ ] Write `manual calculation withholds until measured`. A `calcMode="manual"` workbook's formula rules are unresolved without `formula:calc-manual-saved-values` and follow the saved values with it.
- [ ] Write `one clock per preparation on every leg (S5-M)`. With an injected `now` inside `CfEvaluationContext`, static, text, resident and worker preparations give the same `TODAY()` outcome and the same `clockGen`. A midnight tick invalidates every tile of a `TODAY()`-rule sheet once, and a sheet without clock-dependent rules sees no tick (zero preparations). `NOW()` stays refused.
- [ ] Write `static, print, export, resident, worker and forced-paged agree for paint, a11y and copy/find (Lesson 13)`. On the Stage 5 fixtures, the cross-sheet fixture included, each leg's per-cell properties, the `describeCfWin` descriptor and the copy/find text are pairwise equal. A leg without `CfWorkbookSource` withholds the cross-sheet cells identically on all three surfaces.
- [ ] Write `formula descriptors carry no formula text (S5-Q, Lesson 10)`. A matching expression's accessible descriptor equals the Stage 1 DXF descriptor for the same DXF, is ≤ 128 units, and contains no `=`, reference or function name. Focus order is unchanged.
- [ ] Write `blank cells match formula rules in every view`. `=ISBLANK(B2)` over an unstored B2 on a CF-only sheet paints on static, print and interactive legs, and hit testing reaches it (the Stage 1 `B2`/`Z5` pattern).
- [ ] Extend `tests/browser/xlsx-cf.spec.ts` with `Stage 5 formula rules paint and announce on every leg, and legs agree`, `Stage 5 cross-sheet edit repaints the dependent sheet`, `Stage 5 formula rules under frozen panes` and `Stage 5 print capture paints formula rules`.
- [ ] Run `npx vitest run tests/xlsx-cf-formula-views.test.ts tests/xlsx-cf-print.test.ts tests/xlsx-cf-export.test.ts tests/xlsx-cf-copy-find.test.ts tests/xlsx-cf-paged-a11y.test.ts tests/xlsx-cf-presentation.test.tsx`, then `npx playwright test --config playwright.viewer.config.ts tests/browser/xlsx-cf.spec.ts --project chromium`, tsc and the eager gates. Commit (`feat(xlsx-cf): one policy and clock for CF formulas in every view`), report, review.

### Task 39: Stage 5 validation, five-family closure gates, resource audit and PR evidence (Claude Sonnet; review per table)

**Files:** Create `tests/xlsx-cf-stage5.test.ts`. Extend `tests/xlsx-cf-oracle-conformance.test.ts` (Stage 5 counts) and `tests/browser/xlsx-cf.spec.ts`. Update `docs/validation/xlsx-cf.md`, the manifest statuses, `scripts/eager-baseline.json` (recorded growth), `src/xlsx/cf/verifiedCapabilities.ts` (via `--write --publish-stage5` only) and `CHANGELOG.md`. Fix product regressions in separate focused commits, each with RED evidence and a different-family re-review.

**Interfaces:**
- **Consumes:** everything above, `auditStage5Manifest`, the Stage 1–4 audits.
- **Produces:** a Stage 5 audit row per case and variant (status, evidence kind, Excel vs production reading, mode, stamp, byte peak, work units); the generated `VERIFIED_STAGE5_CAPABILITIES`; the PR 5 evidence.

- [ ] Write `every enabled formula capability has verified evidence` and `every unresolved formula shape refuses atomically on both paths`, covering the `formula:*` rows and the formula rows of `scale:`/`bar:`/`icon:` only. RED, then GREEN.
- [ ] **Corpus gate (not vacuous).** Write `formula CF files parse identically on both paths` over the inventoried corpus files and, unconditionally, over `excel-authored-formulas.xlsx` and every Stage 5 saved workbook. The file count must equal the inventory's; a count of 0 corpus files is recorded, never treated as a pass on its own.
- [ ] **Dual gate (Lessons 1, 4).** Run `bun scripts/cf-refs/audit.ts --stage 5` until it reports `stageComplete=true`, with every disposition in {`verified`, `unverified-nondiscriminating`, `refused-unsupported`, `supported-structural`} and none `unaccounted`, AND `npx vitest run tests/xlsx-cf-oracle-conformance.test.ts` green with no injection. Then run `--write --publish-stage5` twice; `git diff --exit-code src/xlsx/cf/verifiedCapabilities.ts` after the second proves idempotency.
- [ ] **Five-family closure.** Re-run the Stage 1–4 audits and conformance under the published Stage 5 set: `audit.ts --stage 1..4` stay ok, and `git diff main -- src/xlsx/cf/verifiedCapabilities.ts` touches only `VERIFIED_STAGE5_CAPABILITIES` and the combined set.
- [ ] Run `npx tsc --noEmit -p .`, `npx tsc --noEmit -p tsconfig.scripts.json`, `npx vitest run`, `bun run gates:core` (including `OV_MODULE_MAP=1 bun run build`), `bun scripts/check-eager.ts`, `bun scripts/eager-budget.ts` and `git diff --check`. Record the summaries, the eager gzip (strictly below 38,720 B), the `eager-baseline.json` delta with its justification, `indexModuleSetHash`, and the `cf` chunk delta against PR 4.
- [ ] Browser: `npx playwright test --config playwright.viewer.config.ts tests/browser/xlsx-cf.spec.ts --project chromium`, plus `--project webkit` if installed. Cover saved, interactive, resident and forced-paged views, the six zoom/DPR pairs, frozen panes, a live cross-sheet edit, a midnight tick (injected clock), print capture, and the formula descriptor. An unavailable engine is a named blocked gate.
- [ ] Run `bun scripts/cf-compare.ts --stage 5 --mode static|interactive|paged|print|export`, bound to the workbook and oracle hashes. Unverified fidelity is labelled `*unverified*`.
- [ ] Run `bun scripts/cf-refs/audit.ts --stage 5` resource scenarios:
  - a 1M-row synthetic sheet with 100 formula rules over `A:A` sharing one snapshot;
  - `SUM($A:$A)` and `COUNTIF($A:$A,A2)>1` on 5,000 visible cells over a declared 1M-row `A:A` with 10,000 populated cells: hoisted, clamped to populated bounds, one chunked scan per rule paid by the scan budget, per-cell calls ≤ 128 units; the same rules over a column with about a million **populated** rows refuse at the scan budget with one diagnostic, as Stage 1–4 statistics do; and a relative-range `SUM(A1:A200)` over the same cells is refused before any `evaluateNode` call;
  - snapshot pressure at the cap and cap+1;
  - 16 and 17 referenced sheets;
  - a refused footprint beside a ready rule;
  - 20 open/read/cancel/close cycles, including close during a cross-sheet pending read.

  Pass when there are zero unowned live bytes at every transition, and zero formula compiles, snapshot reads or clock reads on a no-formula baseline.
- [ ] Publish one row per Stage 5 `[measure]` item in `docs/validation/xlsx-cf.md`: ruling, capability, evidence kind, status and the exact refusal still in force, including every re-deferral of the deferral sweep.
- [ ] Update the changelog. The reviewer verdict is recorded in `.superpowers/sdd/<plan>/task-39-review.md` (APPROVE), then the whole-branch high-effort review by the other family. Budget one extra PR-bot round for cross-path semantics. Prepare the PR 5 evidence and leave the merge to the owner.

### Lesson-to-task mapping

| Lesson (brief numbering) | Task | Test |
|---|---|---|
| 1 Oracle first + production conformance | 31, 35, 39 | `stage 5 formula cells are withheld until their evidence lands`; `expression rule paints exactly the DisplayFormat cells (Lesson 1)`; the dual gate |
| 2 Exact or withheld | 32–38 | `compile admits the allow-list and refuses every out-of-scope node statically`; `truth readings are gated`; `a read outside the snapshot taints and discards` |
| 3 Non-discriminating ⇒ withhold, ≥ 2 cells | 31 | `semantic DisplayFormat evidence needs two decisive cells and a control (Lesson 3)` |
| 4 Publication only through the audit; dual gate; matched reading = production reading | 31, 39 | `audit --stage 5 --write is idempotent`; `a publication requires production's reading to equal Excel's matched reading`; Task 39 dual gate |
| 5 Whole-sheet invalidation | 36, 38 | `remote cache entries follow their reply's dependencies` (every tile of Sheet1); the midnight-tick test |
| 6 Measurer bytes; finite near ±1e308 | 32, 34, 36, 37 | `formula definitions are priced`; `formula results stay finite near ±1e308`; `dependency bytes are measurer-derived`; `formula bounds stay finite` |
| 7 Eager size; CF out of eager/paged-worker | 32, 36, 39 | lazy file list + `indexModuleSetHash` assertion; recorded `canRuleMatchBlanks`/stamp deltas |
| 8 Linux CI: skipIf + TZ pin | 31, 38 | osacompile under `it.skipIf(!hasOsacompile)`; TZ-pinned clock tests |
| 9 Deterministic budgets; fresh fallback budget | 34, 39 | `each synchronous call stays within one checkpoint and the loop yields between cells`; `traversal work is charged per occurrence before it happens`; `text work is bounded by measured payload before invocation`; active-units extension |
| 10 A11y from semantics; legs announce identically | 38 | `formula descriptors carry no formula text`; the every-leg agreement test |
| 11 Coarse plans: decorations only on a painted host | 38 | inherited unchanged: formula rules add no decoration; the agreement test includes a coarse frame |
| 12 Excel canonicalises; Excel-authored source; record strips | 31, 37 | `stage5-formula-authored`; one shape per workbook; `x14 colour scales follow the recorded dialect`; `activePresent follows its recorded outcome` |
| 13 Every view agrees; missing inputs fail closed identically | 36, 38 | `static, print, export, resident, worker and forced-paged agree…`; `every leg reads one value policy`; `cross-sheet edge cases refuse identically` |
| 14 No post-hoc fitting | 31, 39 | hypotheses fixed in each case before capture; reviewer checklist item (Review Focus 10) |
| 15 Charge before allocating (growth peaks, decoded caches) | 34, 36 | translated-tree reservation; snapshot growth charges old + new; cross-sheet decoded pages admitted under the batch lease |

### Stage 5 owner decisions

1. **OD-11 — splitting.** Decided 2026-11-16: Stage 5 ships as one PR, Tasks 31–39 in order (owner decision).
   - Alternative: three sub-PRs (same-sheet, cross-sheet, cfvo/x14/closure). Cost: a group-scoped audit and three closure runs.
2. **OD-12 — calculation policy.** Decided 2026-11-16: option (a), CF formulas read the grid-displayed value of every cell (cached or engine, whichever the grid shows), never recompute, still evaluate under `'cached'` (a CF rule has no cache), treat an unavailable cached value as unresolved, and withhold manual-mode workbooks until `formula:calc-manual-saved-values` (owner accepted the recommendation).
   - Alternative (b): always read the saved `<v>`. Cost: CF can disagree with the grid the user sees.
3. **OD-13 — function coverage.** Decided 2026-11-16: the closed tier-1 allow-list of S5-B (only functions registered on main), with no engine additions in Stage 5 (owner accepted the recommendation).
   - Alternative: admit the lookup family. Cost: lookup semantics in CF need their own discriminating cases.
   - Alternative: implement `SEARCH`/`FIND`/`ISEVEN`/`ISODD` in the shared engine. Cost: changes worksheet formulas and needs its own engine oracle.
4. **OD-14 — `NOW()` and sub-day volatility.** Decided 2026-11-16: `NOW()` is withheld, and `TODAY()` is admitted under the Stage 1 clock and the M2 midnight tick (owner accepted the recommendation).
   - Alternative: a periodic CF clock tick for `NOW()` sheets. Cost: a new timer and an invalidation path in every view.
5. **OD-15 — x14 colour scales (OD-3 pickup).** Decided 2026-11-16: in scope only if Excel's dialog writes the x14 form for a cross-sheet bound; otherwise the outcome is recorded and the shape re-deferred (owner accepted the recommendation).
   - Alternative: ship from hand-authored x14 XML. Cost: a dialect Excel may never write, so there is no evidence.
6. **OD-16 — `activePresent`.** Decided 2026-11-16: a permanent recorded refusal is accepted if Excel cannot author it or strips it on resave (owner accepted the recommendation).
   - Alternative: ship from hand-authored XML that Excel never keeps. Cost: unverifiable semantics.

### Self-review

- **Coverage.**
  - §4 row 5 and §10 item 5 → Goal, deferral sweep, Tasks 32–37.
  - §5.1 non-literal operands → S5-F, Task 35. §5.2/§5.3 divergent formulas → S5-G, Task 35.
  - §5.7 composition → S5-O, S5-STUB, Task 35. §5.8–§5.10 formula bounds → S5-H, Task 37.
  - §5.11 anchor, truthiness, `activePresent`, clock, calc policy, cross-sheet → S5-C/D/J/M/L/K, Tasks 34, 36–38.
  - §6.3 paged dependency preparation and budgets (including line 245's `SUM(A:A)` repeat) → S5-K/N/HOIST, Tasks 34, 36. §6.5/§6.6 → S5-K stale-at-commit, Task 36.
  - §6.7 → S5-Q, Task 38. §6.8 → S5-R, Task 36. §7 → Global Constraints delta, caps, Task 39. §8 → S5-A, Task 31. §9 items 5/8 → Review Focus 1, 4, 5.
- **Interfaces.** Every Consumes and architecture citation was printed on `fc61c6c6` after drafting (a mechanical pass over every `file:line`). Every Produces name was grepped absent with `rtk proxy grep`. Corrected facts: `translateSharedFormula` lives at `shared.ts:386`, not `workbook.ts:938`; the Stage 1 "external source change" coordinator does not exist on main; `ResolvedSemantics` fields are `timeZone`/`epochNowMs`.
- **Main-branch facts that change the plan.**
  - Conditional aggregates and whole-column references need `ReferenceServices`, so the adapter builds them over its snapshot.
  - `SEARCH`/`FIND` are unregistered, so Excel's own divergent text formulas stay withheld.
  - Every CF source and stamp is single-sheet, so cross-sheet needs `dependencyGen` and a workbook seam.
  - The Stage 2–4 rejected twins lost all CF, so every malformed shape gets its own workbook.
  - `canRuleMatchBlanks` is eager and keyed on the unresolved `expression`.
  - Paint and text default to different value policies.
- **Round-2b (S5-HOIST).** Absolute-range aggregates and the `COUNTIF($A:$A,A2)` idiom are hoisted into the chunked statistics machinery, so the 128-unit per-call cap no longer refuses them; the hoist is exact-or-refused against the engine's own `COUNTIF`.
- **Proportion.** Nine tasks against Stage 4's seven (Task 33 was added in round 1 for the `timePeriod` fail-open). Three things are new compared with Stage 4: the formula adapter and snapshot (Task 34), the cross-sheet generation (Task 36), and the reused-capture stub successors (Tasks 31, 35).

## Self-review checklist for this plan

- [x] Independent plan review confirms Stage 1 owns composition, invalidation and every-view foundations; later stages only add rule families.
- [x] Every Stage 1 task has an implementer, a qualifying different-family reviewer, exact files/contracts, RED/GREEN test names and concrete boundary inputs.
- [x] Every Stage 1 semantic uncertainty has an oracle case ID or remains refused; no case is marked measured by plan authorship.
- [x] Resource ceilings are local CF limits under existing profile budgets; admission precedes allocation and live ownership survives cancellation.
- [x] Spec numeric constraints, source facts and no-CF behavior are checked; no implementation checkbox is ticked in advance.

Self-review completed 2026-11-06: checked the approved spec at `e37d6179`, source seams and exact browser config. The writing-plans skill advisory reviewer returned Approved; it was a GPT advisory review, not the required different-family implementation gate. All implementation task boxes remain unchecked.
