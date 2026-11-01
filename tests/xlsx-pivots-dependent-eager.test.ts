import { readFileSync, statSync } from 'node:fs'
import { dirname, extname, resolve } from 'node:path'
import ts from 'typescript'
import { expect, it } from 'vitest'

function eagerFormulaGraph(entry: string): Set<string> {
  const seen = new Set<string>()
  const visit = (file: string): void => {
    const absolute = resolve(file)
    if (seen.has(absolute)) return
    seen.add(absolute)
    const source = ts.createSourceFile(absolute, readFileSync(absolute, 'utf8'), ts.ScriptTarget.Latest, true)
    const specifiers: string[] = []
    for (const statement of source.statements) {
      if (!ts.isImportDeclaration(statement) && !ts.isExportDeclaration(statement)) continue
      const clause = ts.isImportDeclaration(statement) ? statement.importClause : undefined
      if (clause?.isTypeOnly) continue
      if (ts.isExportDeclaration(statement) && statement.isTypeOnly) continue
      if (clause?.namedBindings && ts.isNamedImports(clause.namedBindings) && clause.namedBindings.elements.length && clause.namedBindings.elements.every(element => element.isTypeOnly)) continue
      const specifier = statement.moduleSpecifier
      if (specifier && ts.isStringLiteral(specifier) && specifier.text.startsWith('.')) specifiers.push(specifier.text)
    }
    for (const specifier of specifiers) {
      const target = resolve(dirname(absolute), specifier)
      const candidates = extname(target) ? [target] : [`${target}.ts`, `${target}.tsx`, resolve(target, 'index.ts'), resolve(target, 'index.tsx')]
      const found = candidates.find(candidate => { try { return statSync(candidate).isFile() } catch { return false } })
      if (found) visit(found)
    }
  }
  visit(entry)
  return seen
}

it('the static eager formula graph has no pivot implementation imports (no build artifact needed)', () => {
  const files = eagerFormulaGraph(resolve(process.cwd(), 'src/xlsx/formula/index.ts'))
  const pivotModules = [...files].map(file => file.replace(/\\/g, '/')).filter(file => /\/src\/xlsx\/pivots\/(?!stub\.ts$)/.test(file))
  expect(pivotModules).toEqual([])
})
