/**
 * Pure helpers for the §8.1 eager gates (`scripts/core-gates.ts`) and the
 * Stage-0 T11 probe. Operate on the module maps `scripts/vite-module-map.ts`
 * writes; no build or file access here.
 */
import { createHash } from 'node:crypto'
import type { ModuleMap } from './eager-graph'

/** Chunk files reachable from `entry` through static imports (entry
 * included), sorted. Dynamic imports are lazy and never followed. */
export function staticClosure(map: ModuleMap, entry: string): string[] {
  const seen = new Set<string>()
  const queue = [entry]
  while (queue.length > 0) {
    const file = queue.pop()!
    if (seen.has(file) || !map[file]) continue
    seen.add(file)
    queue.push(...map[file].imports)
  }
  return [...seen].sort()
}

/** Source trees that hold parsers and painters: the stream client must not
 * statically reach any of them (§8.1 gate 1). */
export const ENGINE_PREFIXES = [
  'src/xlsx/', 'src/docx/', 'src/pptx/', 'src/drawing/', 'src/render/', 'src/charts/',
  'src/slideshow/', 'src/odf/', 'src/odt/', 'src/rtf/', 'src/components/',
] as const

/** Repo-relative ids of engine modules bundled into `chunks`, sorted. */
export function engineModulesIn(map: ModuleMap, chunks: readonly string[], root: string): string[] {
  const out = new Set<string>()
  const prefix = root.replace(/\\/g, '/')
  for (const chunk of chunks) {
    for (const id of map[chunk]?.modules ?? []) {
      const norm = id.replace(/\\/g, '/')
      const rel = norm.startsWith(prefix) ? norm.slice(prefix.length) : norm
      if (ENGINE_PREFIXES.some((p) => rel.startsWith(p))) out.add(rel)
    }
  }
  return [...out].sort()
}

const REACT = /^react(?:-dom)?(?:\/|$)/

/** Splits built JS into code with string literals replaced by `\u0000<n>\u0000`
 * placeholders and comments removed. Templates are treated as one literal;
 * regex literals are not special-cased (built output rarely needs it, and a
 * miss can only produce a false positive, never hide a violation). */
function scan(source: string): { code: string; strings: string[] } {
  let code = ''
  const strings: string[] = []
  let i = 0
  while (i < source.length) {
    const c = source[i]!
    const next = source[i + 1]
    if (c === '/' && next === '/') {
      const end = source.indexOf('\n', i)
      i = end === -1 ? source.length : end
      continue
    }
    if (c === '/' && next === '*') {
      const end = source.indexOf('*/', i + 2)
      i = end === -1 ? source.length : end + 2
      code += ' '
      continue
    }
    if (c === '"' || c === "'" || c === '`') {
      let j = i + 1
      let value = ''
      while (j < source.length && source[j] !== c) {
        if (source[j] === '\\') { value += source[j + 1] ?? ''; j += 2; continue }
        value += source[j]
        j++
      }
      code += `\u0000${strings.length}\u0000`
      strings.push(value)
      i = j + 1
      continue
    }
    code += c
    i++
  }
  return { code, strings }
}

/** React value imports and `document` global uses in built JS (§8.1 gate 2).
 * Type-only imports are erased by the build, so every hit is real. */
export function purityViolations(source: string): string[] {
  const { code, strings } = scan(source)
  const out: string[] = []
  const lit = '\\u0000(\\d+)\\u0000'
  for (const re of [new RegExp(`\\bfrom\\s*${lit}`, 'g'), new RegExp(`\\bimport\\s*\\(\\s*${lit}`, 'g'), new RegExp(`\\bimport\\s*${lit}`, 'g')]) {
    for (const m of code.matchAll(re)) {
      const spec = strings[Number(m[1])]!
      if (REACT.test(spec)) out.push(`import "${spec}"`)
    }
  }
  if (/(?<![.\w$])document(?![\w$])/.test(code)) out.push('document global')
  return out
}

/** Repo-relative ids of every module in `entry`'s static closure, sorted —
 * the §8.1 gate-3 module-map identity (portable across checkouts; chunk file
 * names carry content hashes, so they are not part of it). */
export function closureModules(map: ModuleMap, entry: string, root: string): string[] {
  const prefix = root.replace(/\\/g, '/')
  const out = new Set<string>()
  for (const chunk of staticClosure(map, entry)) {
    for (const id of map[chunk]!.modules) {
      // Rollup virtual ids (commonjs wrappers) start with NUL; keep them
      // distinguishable but path-relative.
      const virtual = id.startsWith('\u0000')
      const norm = (virtual ? id.slice(1) : id).replace(/\\/g, '/')
      const rel = norm.startsWith(prefix) ? norm.slice(prefix.length) : norm
      out.add(virtual ? `virtual:${rel}` : rel)
    }
  }
  return [...out].sort()
}

/** sha-256 hex of a sorted module-id list. */
export function moduleSetHash(ids: readonly string[]): string {
  return createHash('sha256').update(JSON.stringify([...ids].sort())).digest('hex')
}

/** Added/removed ids between a baseline module list and the current one. */
export function moduleDiff(baseline: readonly string[], current: readonly string[]): { added: string[]; removed: string[] } {
  const b = new Set(baseline)
  const c = new Set(current)
  return { added: current.filter((id) => !b.has(id)), removed: baseline.filter((id) => !c.has(id)) }
}
