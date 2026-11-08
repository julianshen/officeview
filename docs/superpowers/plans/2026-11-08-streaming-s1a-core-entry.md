# Streaming Stage 1a — `officeview/core` entry + eager gates — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship a React-free `officeview/core` entry built in its own Vite pass, the `StreamSource` type with its misuse guard, and the three §8.1 eager gates as a runnable script.

**Architecture:** `src/core-entry.ts` exports only stream-source types, the guard, and a lazy `loadStreamRuntime()` that dynamic-imports `src/stream/index.ts` (empty runtime placeholder in 1a; 1b/1d fill it). A second Vite config (`vite.core.config.ts`) builds `dist/core.js` without wiping `dist/`, emitting `dist/.module-map.core.json`. `scripts/core-gates.ts` checks the core closure ceiling + engine-free closure, purity, and main-entry no-growth against a committed baseline.

**Tech Stack:** TypeScript, Vite 7 lib mode (rollup), Bun scripts, Vitest.

**Spec:** `docs/superpowers/specs/2026-11-07-server-streaming-design.md` — §5 (client layout), §6.1 (`StreamSource`, `HttpSource.kind?: never`, `StreamSourceMisuseError`), §8.1 (gates, as ruled in §2 item 14b), §17 row 1.

**Sub-plan context:** Stage 1 = 1a (this) → 1b recorder/replay/OVSD → 1c server → 1d client integration. 1a touches `src/core/stream.ts`, `src/index.ts`, new `src/core-entry.ts`, new `src/stream/index.ts`, build config and `package.json`. Overlap with the live CF worktree (`.worktrees/cf`, `src/xlsx/**`): none expected — check `git diff main --stat` of `feature/xlsx-cf` before merge.

## Global Constraints

- Eager ceiling: **38,720 B** gzip (Node `zlib.gzipSync`, default level) — same constant as `scripts/eager-budget.ts:EAGER_BUDGET`.
- Main-entry closure baseline: **417,436 B** gzip over `dist/index.js`'s static closure; allowed growth **≤ 16 B**, any growth printed as a diff.
- Main-entry module-map baseline hash: **`a0669339ada2a01085db11e6ebc77fb8914ff42ff6ee16b19c6437a919674dc3`** (Stage-0 T11 `moduleMapHash`, computed as in `scripts/stream-s0/t11-eager.ts`).
- `eager-budget.ts` and `check-eager.ts` keep working unchanged (the main entry keeps its `index.js`-alone check).
- React stays a peer dependency; `dist/core.js` must contain **no** import of `react`/`react-dom`/`react/*` and no reference to the `document` global.
- `wss://` URLs must never reach `fetch` (§6.1).
- Product behaviour of every existing entry is unchanged; the full `bun run test` suite stays green.

## Review Focus

1. **A `StreamSource` passed where a `ByteSource` is expected** (JS callers, `as any`) — must throw `StreamSourceMisuseError` before any network call, never fetch the `wss://` URL. Test in Task 1.
2. **`HttpSource` objects created by existing callers** (no `kind`) — must keep working exactly as before. Test in Task 1.
3. **The second Vite pass deleting the first pass's output** (`emptyOutDir` default true) — `dist/index.js` and `dist/worker.js` must survive. Test in Task 2.
4. **A future 1b/1d import from `core-entry.ts` into the engine** (e.g. importing a painter for a type) — the closure gate must name the offending module, not just fail on bytes. Test in Task 3.
5. **Type-only imports from React** in core modules (`import type { … } from 'react'`) — erased, so allowed; a value import must fail purity. Test in Task 3.

---

### Task 1: `StreamSource`, `HttpSource.kind?: never`, `StreamSourceMisuseError`

**Files:**
- Modify: `src/core/stream.ts` (`HttpSource` at :42–46; `readSource` :207; `openByteStream` :368)
- Create: `src/stream/source.ts`
- Test: `tests/stream/source.test.ts`

**Interfaces:**
- Produces (in `src/stream/source.ts`, React-free, no engine imports):
  - `export interface StreamSource { kind: 'stream'; url: string; token?: string }` (doc comments verbatim from spec §6.1)
  - `export function isStreamSource(v: unknown): v is StreamSource` — true iff non-null object with `kind === 'stream'` and string `url`
  - `export class StreamSourceMisuseError extends Error { readonly name = 'StreamSourceMisuseError' }` — message `'A StreamSource cannot be read as bytes; pass it to <OfficeFile data> to open a stream session.'`
- `HttpSource` gains `kind?: never`.

- [ ] **Step 1: Write the failing tests**

```ts
it('readSource refuses a StreamSource without fetching', async () => {
  const fetchSpy = vi.spyOn(globalThis, 'fetch')
  await expect(readSource({ kind: 'stream', url: 'wss://x/s' } as never)).rejects.toBeInstanceOf(StreamSourceMisuseError)
  await expect(openByteStream({ kind: 'stream', url: 'wss://x/s' } as never)).rejects.toBeInstanceOf(StreamSourceMisuseError)
  expect(fetchSpy).not.toHaveBeenCalled()
})
it('a Promise resolving to a StreamSource is refused too', async () => { /* same, wrapped in Promise.resolve */ })
it('HttpSource without kind still downloads', async () => { /* stub fetch → Response(new Uint8Array([1,2,3])); readSource({url:'https://x/a.docx'}) resolves 3 bytes */ })
it('isStreamSource', () => {
  expect(isStreamSource({ kind: 'stream', url: 'wss://a' })).toBe(true)
  expect(isStreamSource({ url: 'wss://a' })).toBe(false)
  expect(isStreamSource({ kind: 'stream' })).toBe(false)
  expect(isStreamSource(null)).toBe(false)
})
```
Type-level: `// @ts-expect-error` on `const h: HttpSource = { kind: 'stream', url: 'wss://a' }`.

- [ ] **Step 2: Run** `bun run test tests/stream/source.test.ts` — FAIL (module not found).
- [ ] **Step 3: Implement** `src/stream/source.ts`; in `readSource`/`openByteStream`, after awaiting a promised source, `if (isStreamSource(source)) throw new StreamSourceMisuseError()` as the first statement that inspects the source.
- [ ] **Step 4: Run** the test file, then `bun run test` — all PASS; `npx tsc --noEmit` clean.
- [ ] **Step 5: Commit** `feat(stream): StreamSource type and misuse guard (§6.1)`.

### Task 2: `src/core-entry.ts` + separate Vite pass + package export

**Files:**
- Create: `src/core-entry.ts`, `src/stream/index.ts`, `vite.core.config.ts`
- Modify: `package.json` (`exports`, `scripts.build`), `scripts/vite-module-map.ts` (file-name option), `src/index.ts` (re-export the §6.1 types)
- Test: `tests/build/core-entry.test.ts` (runs against a built `dist/`; skipped with a clear message when `dist/core.js` is absent)

**Interfaces:**
- `src/core-entry.ts` exports exactly: `type StreamSource`, `isStreamSource`, `StreamSourceMisuseError` (from `./stream/source`), and `export function loadStreamRuntime(): Promise<typeof import('./stream/index')>` = `import('./stream/index')`.
- `src/stream/index.ts` (1a placeholder): `export const STREAM_RUNTIME_VERSION = 1`. 1b/1d add the replayer, frame codec and `RemoteOfficeDoc` here.
- `moduleMap(fileName = '.module-map.json')`.
- `src/index.ts` adds `export { isStreamSource, StreamSourceMisuseError } from './stream/source'` and `export type { StreamSource } from './stream/source'` — the only intended main-entry map change (Task 3's identity gate records it).
- `package.json` `exports["./core"] = { "types": "./dist/core-entry.d.ts", "import": "./dist/core.js" }`; `build` becomes `tsc --noEmit && vite build && vite build -c vite.core.config.ts && tsc -p tsconfig.build.json`.

`vite.core.config.ts`: lib entry `{ core: 'src/core-entry.ts' }`, `formats: ['es']`, `build.emptyOutDir: false`, same React `external` regex as `vite.config.ts`, no CSS/notice plugins, `chunkFileNames: 'core-[name]-[hash].js'` (no collision with main-pass chunks), and `moduleMap('.module-map.core.json')` when `OV_MODULE_MAP` is set.

- [ ] **Step 1: Write the failing test**

```ts
it('builds dist/core.js beside the main entries', () => {
  for (const f of ['index.js', 'worker.js', 'core.js', 'core-entry.d.ts']) expect(existsSync(join(DIST, f))).toBe(true)
})
it('core.js exports the stream-source surface only', async () => {
  const mod = await import(pathToFileURL(join(DIST, 'core.js')).href)
  expect(Object.keys(mod).sort()).toEqual(['StreamSourceMisuseError', 'isStreamSource', 'loadStreamRuntime'])
  expect((await mod.loadStreamRuntime()).STREAM_RUNTIME_VERSION).toBe(1)
})
```
- [ ] **Step 2: Run** `OV_MODULE_MAP=1 bun run build && bun run test tests/build/core-entry.test.ts` — FAIL.
- [ ] **Step 3: Implement** the files above.
- [ ] **Step 4: Run** the same command — PASS; `dist/.module-map.json` and `dist/.module-map.core.json` both exist; `bun scripts/eager-budget.ts` and `bun scripts/check-eager.ts` both exit 0.
- [ ] **Step 5: Commit** `feat(core): officeview/core entry in a separate Vite pass (§8.1)`.

### Task 3: `scripts/core-gates.ts` — the three §8.1 gates

**Files:**
- Create: `scripts/core-gates.ts`, `scripts/eager-baseline.json`, `tests/helpers/core-gates.ts` (pure functions)
- Modify: `package.json` (`"gates:core": "OV_MODULE_MAP=1 bun run build && bun scripts/core-gates.ts"`)
- Test: `tests/build/core-gates.test.ts` (pure, synthetic `ModuleMap`s — no build)

**Interfaces (in `tests/helpers/core-gates.ts`, reusing `ModuleMap` from `tests/helpers/eager-graph.ts`):**
- `staticClosure(map: ModuleMap, entry: string): string[]` — chunk files reachable via `imports` only (move `staticChunkClosure` from `scripts/stream-s0/t11-eager.ts` here; t11 imports it).
- `ENGINE_PREFIXES = ['src/xlsx/', 'src/docx/', 'src/pptx/', 'src/drawing/', 'src/render/', 'src/charts/', 'src/slideshow/', 'src/odf/', 'src/odt/', 'src/rtf/', 'src/components/']`
- `engineModulesIn(map: ModuleMap, chunks: string[]): string[]` — module ids (repo-relative) in those chunks that start with an engine prefix.
- `purityViolations(source: string): string[]` — value imports of `/^react(-dom)?(\/|$)/` (static `import … from`, `export … from`, dynamic `import('react…')`) plus any `\bdocument\b` identifier use outside string literals and comments; type-only imports cannot appear in built JS, so a hit is always real.
- `moduleMapHash(map: ModuleMap, entry: string): string` — exactly the t11 formula (sorted modules + sorted imports, sha256 hex).
- `scripts/eager-baseline.json`: `{ "indexClosureGzip": 417436, "indexModuleMapHash": "a0669339…74dc3", "allowedGrowth": 16, "recordedChanges": [] }`.

`scripts/core-gates.ts` prints one line per gate and exits 1 if any fails:
1. `core closure <n> B over <k> chunk(s) ≤ 38720` **and** `engine modules in core closure: none` (else lists them).
2. `core purity: ok` (else lists violations).
3. `index closure <n> B (baseline 417436, Δ <d>)` — fail if Δ > 16. `index map hash` — if it differs from baseline, print added/removed modules and fail **unless** the diff exactly equals an entry in `recordedChanges` (Task 2's `src/stream/source.ts` addition is recorded there in this task, with its Δ bytes).

- [ ] **Step 1: Write failing tests** — closure follows `imports` not `dynamicImports`; an engine module in a statically imported chunk is reported by id; a dynamic-only engine chunk is not; `purityViolations` flags `import{a}from"react"`, `import("react-dom/client")`, `document.createElement`, and ignores `"document"` in a string and `// document` in a comment; `moduleMapHash` is order-independent.
- [ ] **Step 2: Run** `bun run test tests/build/core-gates.test.ts` — FAIL.
- [ ] **Step 3: Implement** the helpers and the script; record Task 2's main-map change in `recordedChanges`.
- [ ] **Step 4: Run** `bun run gates:core` — all three gates print ok; exit 0. Run `bun run s0:napi scripts/stream-s0/t11-eager.ts` still works (imports the moved helper).
- [ ] **Step 5: Commit** `feat(build): §8.1 core gates — closure, purity, main-entry no-growth`.

### Task 4: Acceptance — §17 row 1 "T11 gates green"

- [ ] `bun run gates:core` exit 0; paste its output into `docs/validation/streaming-s1.md` under "T11".
- [ ] `bun run test` full suite green; `npx tsc --noEmit` clean; `bun scripts/eager-budget.ts` ≤ 38,720.
- [ ] Commit `docs(stream-s1): T11 gates green`.
