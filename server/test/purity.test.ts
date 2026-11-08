import { expect, it } from 'bun:test'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { sourceClosure } from './source-graph'

const root = resolve(import.meta.dir, '../..')

it('the server never reaches React, the components layer or the DOM', () => {
  const { files, packages } = sourceClosure(resolve(root, 'server/src/main.ts'), root)
  expect(packages.filter((p) => /^react(-dom)?(\/|$)/.test(p))).toEqual([])
  expect(files.filter((f) => f === 'src/components/OfficeFile.tsx' || f === 'src/components/OfficeDoc.tsx' || f.endsWith('.tsx'))).toEqual([])
  // `document` as a global (not a property/identifier fragment) in server code.
  const serverFiles = files.filter((f) => f.startsWith('server/'))
  const offenders = serverFiles.filter((f) => /(^|[^.\w$])document\s*[.[]/.test(readFileSync(resolve(root, f), 'utf8')))
  expect(offenders).toEqual([])
})
