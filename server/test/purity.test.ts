import { expect, it } from 'bun:test'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { sourceClosure } from './source-graph'
import { purityViolations } from '../../tests/helpers/core-gates'

const root = resolve(import.meta.dir, '../..')

it('the server never reaches React, the components layer or the DOM', () => {
  const { files, packages } = sourceClosure(resolve(root, 'server/src/main.ts'), root)
  expect(packages.filter((p) => /^react(-dom)?(\/|$)/.test(p))).toEqual([])
  expect(files.filter((f) => f === 'src/components/OfficeFile.tsx' || f === 'src/components/OfficeDoc.tsx' || f.endsWith('.tsx'))).toEqual([])
  // No React import or `document` global in server code (strings/comments stripped).
  const offenders = files.filter((f) => f.startsWith('server/')).flatMap((f) => purityViolations(readFileSync(resolve(root, f), 'utf8')).map((v) => `${f}: ${v}`))
  expect(offenders).toEqual([])
  expect(files.length).toBeGreaterThan(100)   // the walk really reached the engine
})
