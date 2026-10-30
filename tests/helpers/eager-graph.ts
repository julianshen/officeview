/**
 * Static-import closure of the built entry, from the module map that
 * scripts/vite-module-map.ts writes. Shared by tests/ui-eager-graph.test.ts
 * and scripts/check-eager.ts.
 */
export type ModuleMap = Record<string, { modules: string[]; imports: string[]; dynamicImports: string[] }>

/** Every module bundled into `entry` or a chunk it statically imports.
 * Dynamic imports are lazy and never followed. */
export function eagerModules(map: ModuleMap, entry = 'index.js'): Set<string> {
  const modules = new Set<string>()
  const seen = new Set<string>()
  const queue = [entry]
  while (queue.length > 0) {
    const file = queue.pop()!
    if (seen.has(file)) continue
    seen.add(file)
    const chunk = map[file]
    if (!chunk) continue
    for (const id of chunk.modules) modules.add(id)
    queue.push(...chunk.imports)
  }
  return modules
}

/** Eager modules matching any rule, sorted. Backslashes are normalised so
 * the rules hold on Windows ids too. */
export function forbiddenEager(map: ModuleMap, rules: readonly RegExp[]): string[] {
  return [...eagerModules(map)].filter((id) => {
    const path = id.replace(/\\/g, '/')
    return rules.some((r) => r.test(path))
  }).sort()
}

// Kept out of the eager entry (plan Global Constraints). Under paged/ only
// metrics.ts may be eager. Rules are tested against /-normalised ids; the
// end anchor allows a ?query suffix.
export const EAGER_FORBIDDEN: readonly RegExp[] = [
  /node_modules\/lucide-react\//,
  /node_modules\/motion\//,
  /node_modules\/framer-motion\//,
  /node_modules\/motion-dom\//,
  /src\/components\/viewerControls\.tsx(?:$|\?)/,
  /src\/components\/Icon\.tsx(?:$|\?)/,
  /src\/components\/motion\//,
  /src\/styles\/tokens\.ts(?:$|\?)/,
  /src\/components\/paged\/(?!metrics\.ts(?:$|\?))/,
  // The chart engine is lazy (spec 6.1); chartLoader, chartModel and the placeholder painter stay eager.
  /src\/charts\//,
  /src\/xlsx\/chart\.ts(?:$|\?)/,
  /src\/xlsx\/tableStyles\//,
  /src\/xlsx\/pivots\/(?!stub\.ts(?:$|\?))/,
  /src\/xlsx\/sparklines\/(?!model\.ts(?:$|\?)|parseGroups\.ts(?:$|\?)|resolve\.ts(?:$|\?))/,
  // The slideshow and the slide appearance enrichment are lazy (slideshow spec 8).
  /src\/slideshow\//,
  /src\/pptx\/appearance\//,
]

// The paged worker's chunks (spec §1.5, §6.2): a large workbook loads the table-style
// built-ins and the styles-extras parser only for a styled table, pivot, slicer or timeline,
// so neither may sit in the static closure of the chunks every paged open loads.
export const PAGED_WORKER_ROOTS: readonly RegExp[] = [
  /src\/xlsx\/view\/pagedSource\.ts(?:$|\?)/,
  /src\/xlsx\/storage\/pagedImport\.ts(?:$|\?)/,
]
export const PAGED_WORKER_FORBIDDEN: readonly RegExp[] = [
  /src\/xlsx\/tableStyles\/builtins(?:\/|\.ts(?:$|\?))/,
  /src\/xlsx\/tableStyles\/parseStylesExtras\.ts(?:$|\?)/,
]

/** Each chunk holding a paged-worker root, with the forbidden modules in its static closure. */
export function pagedWorkerViolations(map: ModuleMap): { roots: string[]; missing: string[]; bad: Array<{ chunk: string; module: string }> } {
  const norm = (id: string) => id.replace(/\\/g, '/')
  const holds = (file: string, rule: RegExp) => map[file]!.modules.some(id => rule.test(norm(id)))
  const roots = Object.keys(map).filter(file => PAGED_WORKER_ROOTS.some(rule => holds(file, rule)))
  const missing = PAGED_WORKER_ROOTS.filter(rule => !roots.some(file => holds(file, rule))).map(rule => rule.source)
  const bad: Array<{ chunk: string; module: string }> = []
  for (const chunk of roots) {
    for (const module of eagerModules(map, chunk)) if (PAGED_WORKER_FORBIDDEN.some(rule => rule.test(norm(module)))) bad.push({ chunk, module })
  }
  return { roots, missing, bad }
}
