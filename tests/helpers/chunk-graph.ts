/**
 * Static + dynamic import graph of a directory of emitted ES chunks (relative
 * `./name.js` specifiers only), for build-output assertions. Stream F item 5:
 * WebKit re-evaluates a module worker's entry when a chunk imports it, so
 * worker graphs must keep the entry chunk import-free and acyclic.
 */
import { readFile } from 'node:fs/promises'
import { join } from 'node:path'

export type ChunkGraph = Map<string, Set<string>>

const SPECIFIER = /(?:\bfrom\s*|\bimport\s*\(\s*|\bimport\s*)["']\.\/([^"']+\.js)["']/g

export async function chunkGraph(dir: string, files: readonly string[]): Promise<ChunkGraph> {
  const graph: ChunkGraph = new Map()
  for (const file of files.filter(name => name.endsWith('.js'))) {
    const code = await readFile(join(dir, file), 'utf8')
    graph.set(file, new Set([...code.matchAll(SPECIFIER)].map(match => match[1])))
  }
  return graph
}

export function importersOf(graph: ChunkGraph, target: string): string[] {
  return [...graph].filter(([file, deps]) => file !== target && deps.has(target)).map(([file]) => file).sort()
}

/** Every import cycle, each reported once as the path that closes it. */
export function chunkCycles(graph: ChunkGraph): string[][] {
  const cycles: string[][] = [], state = new Map<string, 'open' | 'done'>(), stack: string[] = []
  const visit = (file: string) => {
    state.set(file, 'open')
    stack.push(file)
    for (const dep of graph.get(file) ?? []) {
      if (state.get(dep) === 'open') cycles.push([...stack.slice(stack.indexOf(dep)), dep])
      else if (!state.has(dep)) visit(dep)
    }
    stack.pop()
    state.set(file, 'done')
  }
  for (const file of [...graph.keys()].sort()) if (!state.has(file)) visit(file)
  return cycles
}
