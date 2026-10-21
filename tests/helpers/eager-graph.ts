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
  // The slideshow and the slide appearance enrichment are lazy (slideshow spec 8).
  /src\/slideshow\//,
  /src\/pptx\/appearance\//,
]
