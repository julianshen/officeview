// @vitest-environment node
/**
 * Task S1-23: the README's "Slideshow" section documents exactly the public
 * slideshow surface. The props table lists every slideshow prop of
 * `OfficeDocProps` (names containing "present" or "slideShow", internal `__`
 * props excluded) with its declared type, and the error-code table lists
 * exactly the members of `SlideShowErrorCode`. Both are read from the source
 * with the TypeScript parser, so a prop or code added or renamed without the
 * README fails here.
 */
import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import ts from 'typescript'

const ROOT = join(__dirname, '..')
const read = (p: string) => readFileSync(join(ROOT, p), 'utf8')
const readme = read('README.md')
const norm = (t: string) => t.replace(/\s+/g, ' ').trim()

function source(path: string): ts.SourceFile {
  return ts.createSourceFile(path, read(path), ts.ScriptTarget.Latest, true, path.endsWith('x') ? ts.ScriptKind.TSX : ts.ScriptKind.TS)
}
function find<T extends ts.Node>(file: ts.SourceFile, test: (n: ts.Node) => n is T): T[] {
  const out: T[] = []
  const walk = (n: ts.Node) => { if (test(n)) out.push(n); ts.forEachChild(n, walk) }
  walk(file)
  return out
}

/** The README "### Slideshow" section, up to the next heading of the same or a higher level. */
function section(): string {
  const start = readme.indexOf('\n### Slideshow\n')
  expect(start, 'README has a "### Slideshow" section').toBeGreaterThan(0)
  const rest = readme.slice(start + 1)
  const end = rest.slice(1).search(/\n#{1,3} /)
  return end < 0 ? rest : rest.slice(0, end + 1)
}

/** The rows of the first table whose header row starts with `| <first> |`, as unescaped cells. */
function table(first: string): string[][] {
  const lines = section().split('\n')
  const at = lines.findIndex(l => l.startsWith(`| ${first} |`))
  expect(at, `README Slideshow table headed "${first}"`).toBeGreaterThan(0)
  const rows: string[][] = []
  for (const line of lines.slice(at + 2)) {
    if (!line.startsWith('|')) break
    rows.push(line.split(/(?<!\\)\|/).slice(1, -1).map(c => c.trim().replace(/\\\|/g, '|')))
  }
  return rows
}
const code = (cell: string) => {
  const m = /^`([^`]+)`$/.exec(cell)
  expect(m, `cell is one code span: ${cell}`).not.toBeNull()
  return m![1]
}

/** Type aliases declared in present.ts, expanded so the README can show the literal union. */
function aliases(): Map<string, string> {
  const file = source('src/components/present.ts')
  return new Map(find(file, ts.isTypeAliasDeclaration).map(a => [a.name.text, norm(a.type.getText(file))]))
}

function slideshowProps(): Map<string, string> {
  const file = source('src/components/OfficeDoc.tsx')
  const props = find(file, ts.isInterfaceDeclaration).find(i => i.name.text === 'OfficeDocProps')
  expect(props, 'OfficeDocProps interface').toBeDefined()
  const known = aliases()
  const out = new Map<string, string>()
  for (const m of props!.members) {
    if (!ts.isPropertySignature(m) || !m.type) continue
    const name = m.name.getText(file)
    if (name.startsWith('__') || !/present|slideShow/i.test(name)) continue
    const type = norm(m.type.getText(file))
    out.set(name, known.get(type) ?? type)
  }
  return out
}

function errorCodes(): string[] {
  const file = source('src/components/present.ts')
  const alias = find(file, ts.isTypeAliasDeclaration).find(a => a.name.text === 'SlideShowErrorCode')
  expect(alias && ts.isUnionTypeNode(alias.type), 'SlideShowErrorCode is a union').toBe(true)
  return (alias!.type as ts.UnionTypeNode).types.map(t => {
    expect(ts.isLiteralTypeNode(t) && ts.isStringLiteral(t.literal), t.getText(file)).toBe(true)
    return ((t as ts.LiteralTypeNode).literal as ts.StringLiteral).text
  })
}

describe('README Slideshow section', () => {
  it('README slideshow props and error codes match the exported types', () => {
    // Props: the same names in both directions, each with its declared type.
    const declared = slideshowProps()
    expect(declared.size, 'slideshow props found in OfficeDocProps').toBeGreaterThanOrEqual(7)
    const documented = new Map(table('Prop').map(([name, type]) => [code(name), norm(code(type))]))
    expect([...documented.keys()].sort()).toEqual([...declared.keys()].sort())
    for (const [name, type] of declared) expect(documented.get(name), `README type of ${name}`).toBe(type)

    // Error codes: exactly the SlideShowErrorCode members, each once.
    const codes = errorCodes()
    const listed = table('Code').map(([c]) => code(c))
    expect(new Set(listed).size).toBe(listed.length)
    expect([...listed].sort()).toEqual([...codes].sort())

    // Every slideshow type the section names is exported from the package entry.
    const entry = read('src/index.ts')
    for (const name of ['OfficeViewerHandle', 'SlideShowSession', 'SlideShowError', 'SlideShowDiagnostic']) {
      expect(section(), `README names ${name}`).toContain(name)
      expect(entry, `src/index.ts exports ${name}`).toMatch(new RegExp(`export (type )?\\{[^}]*\\b${name}\\b[^}]*\\} from './components/present'`))
    }
  })
})
