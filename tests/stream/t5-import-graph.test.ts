/**
 * Piece 4 Task 9 — the T5 import gate (plan D25; spec §9 item 2, "what
 * reaches the client"). The new playback modules must reach no package
 * module through their static value imports, transitively.
 *
 * The walk starts at the D25 modules (`driver.ts`, `plan-stage.ts`,
 * `plan-source.ts`, `show-bundle.ts`, `show-plan.ts`, `plan-compositor.ts`;
 * a module not yet written — Task 10's source and bundle — is skipped until
 * it lands, then walked). It follows every static import and re-export that
 * survives compilation: `import type` / `export type`, and an import whose
 * every named specifier is an inline `type`, are erased and not followed. A
 * side-effect import (`import './x'`) is followed. Dynamic `import()` is not
 * static and is not followed (the lazy chunk boundary). Bare package
 * specifiers (`three`, …) are leaves.
 *
 * The forbidden pattern (controller ruling, 2026-11-15, on the brief's):
 *
 *   src/(pptx/ | core/package | core/zip
 *       | slideshow/(prepare|read)/
 *       | slideshow/render/(raster|plan|direct|scene|backend).ts
 *       | slideshow/morph/match
 *       | slideshow/player/(stage|session|text-builds|hit-test).ts)
 *
 * - `render/plan` is anchored to the file, so `render/plan-compositor` (a D25
 *   module itself) does not match.
 * - `slideshow/read/limits` is allowed: pure reader constants
 *   (`READER_LIMITS`), reached through `engine/behaviours` → `formula` and
 *   through `show-plan`.
 * - `core/zip` (`OfficePackage`, the package itself) is added to the brief's
 *   list; `core/package` matches only `core/package*.ts` modules.
 */
import { existsSync, readFileSync } from 'node:fs'
import { dirname, relative, resolve } from 'node:path'
import ts from 'typescript'
import { describe, expect, it } from 'vitest'

const ROOT = resolve(__dirname, '../..')

/** D25: the new playback modules. */
const ENTRY = [
  'src/slideshow/player/driver.ts',
  'src/stream/remote/plan-stage.ts',
  'src/stream/remote/plan-source.ts',
  'src/stream/remote/show-bundle.ts',
  'src/stream/show-plan.ts',
  'src/slideshow/render/plan-compositor.ts',
]

export const FORBIDDEN = /(^|\/)src\/(pptx\/|core\/(package|zip)|slideshow\/(prepare|read)\/|slideshow\/render\/(raster|plan|direct|scene|backend)\.tsx?$|slideshow\/morph\/match\.tsx?$|slideshow\/player\/(stage|session|text-builds|hit-test)\.tsx?$)/
export const ALLOWED = /(^|\/)src\/slideshow\/read\/limits\.ts$/

/** The value specifiers of one module (type-only imports and re-exports erased). */
export function valueSpecifiers(file: string): string[] {
  const src = ts.createSourceFile(file, readFileSync(file, 'utf8'), ts.ScriptTarget.Latest, false, file.endsWith('x') ? ts.ScriptKind.TSX : ts.ScriptKind.TS)
  const out: string[] = []
  for (const st of src.statements) {
    if (ts.isImportDeclaration(st)) {
      const c = st.importClause
      if (c?.isTypeOnly) continue
      const named = c?.namedBindings
      // `import { type A, type B } from 'x'` (no default, every specifier a type) is erased.
      if (c && !c.name && named && ts.isNamedImports(named) && named.elements.length > 0 && named.elements.every((e) => e.isTypeOnly)) continue
      out.push((st.moduleSpecifier as ts.StringLiteral).text)
    } else if (ts.isExportDeclaration(st) && st.moduleSpecifier) {
      if (st.isTypeOnly) continue
      const named = st.exportClause
      if (named && ts.isNamedExports(named) && named.elements.length > 0 && named.elements.every((e) => e.isTypeOnly)) continue
      out.push((st.moduleSpecifier as ts.StringLiteral).text)
    } else if (ts.isImportEqualsDeclaration(st) && !st.isTypeOnly && ts.isExternalModuleReference(st.moduleReference)) {
      out.push((st.moduleReference.expression as ts.StringLiteral).text)
    }
  }
  return out
}

function resolveRelative(from: string, spec: string): string {
  const base = resolve(dirname(from), spec)
  for (const c of [base, `${base}.ts`, `${base}.tsx`, `${base}/index.ts`, `${base}/index.tsx`]) {
    if (existsSync(c) && !c.endsWith('/') && (c.endsWith('.ts') || c.endsWith('.tsx'))) return c
  }
  throw new Error(`unresolved import '${spec}' in ${relative(ROOT, from)}`)
}

/** Every module reached from `entries` by static value imports: module → the chain that reached it. */
export function reach(entries: readonly string[]): Map<string, string[]> {
  const seen = new Map<string, string[]>()
  const queue: Array<[string, string[]]> = entries.map((e) => [e, [relative(ROOT, e)]])
  while (queue.length) {
    const [file, chain] = queue.shift()!
    if (seen.has(file)) continue
    seen.set(file, chain)
    for (const spec of valueSpecifiers(file)) {
      if (!spec.startsWith('.')) continue
      const to = resolveRelative(file, spec)
      if (!seen.has(to)) queue.push([to, [...chain, relative(ROOT, to)]])
    }
  }
  return seen
}

describe('T5 import gate (D25)', () => {
  it('the pattern matches the package modules and spares the allowed ones', () => {
    const bad = (p: string) => FORBIDDEN.test(p) && !ALLOWED.test(p)
    for (const p of ['src/pptx/parse.ts', 'src/core/zip.ts', 'src/core/package-sniff.ts', 'src/slideshow/read/store.ts', 'src/slideshow/read/deck.ts', 'src/slideshow/prepare/index.ts',
      'src/slideshow/render/raster.ts', 'src/slideshow/render/plan.ts', 'src/slideshow/render/direct.ts', 'src/slideshow/render/scene.ts', 'src/slideshow/render/backend.ts',
      'src/slideshow/morph/match.ts', 'src/slideshow/player/stage.ts', 'src/slideshow/player/session.ts', 'src/slideshow/player/text-builds.ts', 'src/slideshow/player/hit-test.ts']) {
      expect({ p, bad: bad(p) }).toEqual({ p, bad: true })
    }
    for (const p of ['src/slideshow/render/plan-compositor.ts', 'src/slideshow/read/limits.ts', 'src/slideshow/morph/render.ts', 'src/slideshow/player/driver.ts', 'src/slideshow/render/canvas2d.ts', 'src/slideshow/engine/behaviours.ts']) {
      expect({ p, bad: bad(p) }).toEqual({ p, bad: false })
    }
  })

  it('new playback modules reach no package module', () => {
    const entries = ENTRY.map((e) => resolve(ROOT, e)).filter((e) => existsSync(e))
    // The modules Tasks 3, 4 and 8 wrote are always walked.
    expect(entries.length).toBeGreaterThanOrEqual(4)
    const reached = reach(entries)
    const hits = [...reached].filter(([f]) => { const r = relative(ROOT, f); return FORBIDDEN.test(r) && !ALLOWED.test(r) }).map(([, chain]) => chain.join(' → '))
    expect(hits).toEqual([])
    // The walk is real: it reaches the shared engine and compositor.
    const names = [...reached.keys()].map((f) => relative(ROOT, f))
    expect(names).toEqual(expect.arrayContaining(['src/slideshow/engine/behaviours.ts', 'src/slideshow/render/canvas2d.ts', 'src/slideshow/morph/render.ts']))
  })
})
