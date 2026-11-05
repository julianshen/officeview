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

## Stage 2 stub: colour scales (PR 2)

Refine after PR 1 merges. Oracle first: 2/3 points, percentile/percent, degenerate/inverted bounds, numeric-only participation, min/max, disjoint scope and fill conflict precedence. Measure interpolation and theme/tint; do not assume sRGB without the case. Extend admitted statistics and structured fill output; test all views, blank/merged/filter coordinates, shared statistics budgets and generation invalidation. Formula thresholds and visual activation remain unresolved until Stage 5. Assign alternating implementers and independent reviewers during refinement. Exit: verified scale cases, resource/eager/all-view gates, no lost Stage 1 correctness.

## Stage 3 stub: data bars (PR 3)

Oracle first: paired and standalone x14, 10/90 versus explicit 0/100, autoMin/autoMax, negative/positive/mixed zero ranges, middle/none/automatic axes, same-as-positive flags, border/gradient/direction/showValue and narrow-cell clipping. Refine the structured descriptor and palette ownership, full-path device-grid snapping/tile seams, text suppression and accessible indicator labels. Formula thresholds/activation still refuse. Reuse existing composition, not a second priority mechanism. All-view parity and byte/work gates precede merge.

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
