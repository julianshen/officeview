// Tests and sources resolve files relative to the repository (`resolve(__dirname, 'fixtures/…')`, relative imports):
// a developer's absolute checkout path works on that machine only and fails on CI.
import { readdirSync, readFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { expect, it } from 'vitest'

const ROOT = `${resolve(__dirname, '..')}/`
// Built from parts so this file does not match itself.
const FORBIDDEN = [['', 'Users', ''].join('/'), ['', 'home', 'runner', ''].join('/'), ['.', 'worktrees'].join('')]

function* sourceFiles(dir: string): Generator<string> {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (entry.name === 'node_modules' || entry.name === 'fixtures' || entry.name.startsWith('.')) continue
    const path = join(dir, entry.name)
    if (entry.isDirectory()) yield* sourceFiles(path)
    else if (/\.(c|m)?[jt]sx?$/.test(entry.name)) yield path
  }
}

it('no test or source file names an absolute checkout path', () => {
  const hits: string[] = []
  for (const top of ['tests', 'src']) for (const file of sourceFiles(join(ROOT, top))) {
    readFileSync(file, 'utf8').split('\n').forEach((line, index) => {
      if (FORBIDDEN.some(part => line.includes(part))) hits.push(`${file.slice(ROOT.length)}:${index + 1}`)
    })
  }
  expect(hits).toEqual([])
})
