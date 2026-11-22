/**
 * Every pool dispose under server/test is awaited. `Pool.dispose()` resolves once the pool's workers have closed; a
 * test that drops that promise leaves a Worker alive into the next test file of the same `bun test` process, where
 * (Bun 1.4.5) an awaiting test's Worker messages and `setImmediate` callbacks starve (see fixtures/rejection.ts).
 *
 * A textual check, and simple on purpose. It flags a statement that starts with `<name>.dispose()` (or `?.`), after
 * `;`, `{`, `}` or a line start, where `<name>` is bound by `<name> = new Pool(` in the same file, or is `pool` (the
 * name the `mk()` helpers' pools take). Not flagged, by design: `await x.dispose()`, `return x.dispose()`,
 * `() => x.dispose()` (the hook awaits the returned promise), `const closed = x.dispose()` (awaited later, for a test
 * that asserts while the workers still close) and `Promise.all([x.dispose(), …])`. Limits: a pool under another name
 * built by a helper, a pool reached through a property (`h.pool.dispose()`), or a `const closed = …` that is never
 * awaited are not caught. Non-pool `dispose()` calls (composers, units, tracks: synchronous) are ignored.
 */
import { expect, it } from 'bun:test'
import { readdirSync, readFileSync } from 'node:fs'
import { join, relative } from 'node:path'

const files = (dir: string): string[] =>
  readdirSync(dir, { withFileTypes: true }).flatMap(e => e.isDirectory() ? files(join(dir, e.name)) : /\.tsx?$/.test(e.name) ? [join(dir, e.name)] : [])

export function unawaitedPoolDisposes(src: string): number[] {
  const pools = new Set(['pool', ...[...src.matchAll(/\b(\w+)\s*=\s*new Pool\(/g)].map(m => m[1]!)])
  return [...src.matchAll(/(?:^|[;{}])[ \t]*(\w+)\??\.dispose\(\)/gm)]
    .filter(m => pools.has(m[1]!))
    .map(m => src.slice(0, m.index! + m[0].length).split('\n').length)
}

it('the guard catches a bare pool dispose and lets the awaited forms through', () => {
  const src = [
    'const pool = new Pool({ workers: 1 })',
    '} finally { pool.dispose() }',
    'const roomy = new Pool({ workers: 1 }); roomy.dispose()',
    'composer.dispose()',
    'await pool.dispose()',
    'afterAll(() => pool.dispose())',
    'const closed = pool?.dispose()',
    'await Promise.all([pool.dispose()])',
  ].join('\n')
  expect(unawaitedPoolDisposes(src)).toEqual([2, 3])
})

it('no test file drops a pool dispose', () => {
  const hits = files(import.meta.dir)
    .filter(f => f !== import.meta.path)
    .flatMap(f => unawaitedPoolDisposes(readFileSync(f, 'utf8')).map(line => `${relative(import.meta.dir, f)}:${line}`))
  expect(hits).toEqual([])
})
