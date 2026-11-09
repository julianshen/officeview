# CF Stage 2 Task 13: Scale model, resident and streamed parsing, evidence gates

## Plan & TDD Progression

### Phase 1: Core Model & Evidence Injection
- [x] Test 1: Model types and evidence injection helpers in `src/xlsx/cf/model.ts` and `src/xlsx/cf/evidence.ts`
  - `CfScaleBound`, `CfScalePoint`, resolved `colorScale` on `CfDefinition`
  - Canonical colour-scale write mask `['colorScale', 'fill']` exported
  - `injectCfEvidence` allows tests to inject verified capabilities without modifying `VERIFIED_CAPABILITIES`

### Phase 2: Canonical Write Mask & Two-Scale Withholding
- [x] Test 2: Canonical colour-scale write mask is `['colorScale', 'fill']` for both resolved and unresolved scales
  - Resolved scale carries writes `['colorScale', 'fill']`
  - Unresolved scale carries writes `['colorScale', 'fill']`
  - Proves higher-priority scale's `colorScale` slot withholds a lower scale on the same cell per Ruling S2-I ("two resolved scales on one cell withhold the fill")

### Phase 3: Spec Example Scale & Evidence Gates
- [x] Test 3: Spec example scale resolves to three ordered points
  - With evidence injected (`scale:interpolation-spec-palette`), spec §3 example yields min / percentile 50 / max with points `FFF8696B`, `FFFFEB84`, `FF63BE7B`
  - Without evidence it yields `cf-unsupported-rule` and the scale's fill is tainted
  - Sibling assertion to Stage 1 Task 3 test proves the rule resolves under evidence

### Phase 4: Scale Structure Refusals
- [x] Test 4: Scale structure refusals (S2-C, S2-D, S2-K, OD-3)
  - 2 cfvo + 3 colors (count mismatch or non 2/3) → `cf-invalid-rule`
  - Missing `type` attribute on cfvo → `cf-invalid-cfvo`
  - `num val="$A$1"`, `type="formula"`, `activePresent="1"` → `cf-unsupported-rule`
  - `min val="0"` and mid-list `min` → unresolved without evidence (`scale:thresholds-legacy-min-val-0`, `scale:thresholds-min-midlist-max-first`)
  - 6 cfvo → `cf-resource-limit` before construction
  - Standalone x14 scale, literal or reference `xm:f` → `cf-unsupported-rule` (OD-3, S2-K)
  - Diagnostics deduplicated and sheet cap stays at 64

### Phase 5: Endpoint Colour Resolution
- [x] Test 5: Endpoint colours resolve once on both resident and paged paths
  - Theme 4 with tint -0.4 resolves to `resolveSheetColor` output under the fixture theme, identically on resident and paged paths
  - `auto` and alpha other than `FF` resolve to unresolved without `#000000` fallback

### Phase 6: Streamed and Resident Parity
- [x] Test 6: Resident and streamed CF are deep equal at every chunk boundary
  - Extend chunk boundary test with fixture containing 2-colour, 3-colour, theme and x14 scales
  - All tests in `tests/xlsx-cf-scale-parse.test.ts`, `tests/xlsx-cf-parse.test.ts`, `tests/xlsx-cf-paged-parse.test.ts` pass

### Phase 7: Verification, Commit & Report
- [x] Test 7: All gates pass, commit with trailers, write report, create done marker
  - `npx tsc --noEmit -p tsconfig.json`
  - `npx vitest run tests/xlsx-cf-`
  - `OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts && bun scripts/eager-budget.ts` (eager limit 38720)
  - Commit (`feat(xlsx-cf): parse colour-scale rules behind evidence gates`)
  - Write report to `task-13-report.md`
  - Create `/Users/julianshen/prj/officeview/.worktrees/_prompts/cf-s2-t13.done`
