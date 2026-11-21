/**
 * No `expect(p).rejects` / `.resolves` under server/test: on Bun 1.4.5 they stall while a Worker is alive in the
 * process (see fixtures/rejection.ts), so use `rejection` / `rejectionMessage` / a plain `await` instead.
 */
import { expect, it } from 'bun:test'
import { readdirSync, readFileSync } from 'node:fs'
import { join, relative } from 'node:path'

const files = (dir: string): string[] =>
  readdirSync(dir, { withFileTypes: true }).flatMap(e => e.isDirectory() ? files(join(dir, e.name)) : /\.tsx?$/.test(e.name) ? [join(dir, e.name)] : [])

it('no test file awaits a promise through expect().rejects or .resolves', () => {
  // Whole-file match (\s spans newlines), so a chain split across lines is caught too.
  const hits = files(import.meta.dir)
    .filter(f => !f.endsWith(join('fixtures', 'rejection.ts')) && f !== import.meta.path)
    .flatMap(f => {
      const src = readFileSync(f, 'utf8')
      return [...src.matchAll(/\.\s*(?:rejects|resolves)\s*\./g)].map(m => `${relative(import.meta.dir, f)}:${src.slice(0, m.index).split('\n').length}`)
    })
  expect(hits).toEqual([])
})
