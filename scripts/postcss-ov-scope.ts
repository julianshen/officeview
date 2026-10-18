/**
 * PostCSS stage that scopes officeview's compiled Tailwind output to the theme
 * root (spec §6.8, plan Global Constraints "Scoping stage").
 *
 * It runs through Vite's `css.postcss.plugins`, which is after
 * `@tailwindcss/vite` has generated the CSS and before Vite consumes it, in the
 * library build, the dev server, `?inline` imports and vitest alike. It only
 * touches src/styles/officeview.css (any Vite query stripped) and:
 *  - drops `@property` rules (they register globally);
 *  - unwraps the `@layer properties { @supports … { *, ::before, … } }` --tw-*
 *    fallback into one unlayered, scoped rule placed before the utilities;
 *  - appends `:is([data-ov-theme-root], [data-ov-theme-root] *)` to the
 *    leading class of every `.ov\:` utility selector (before any pseudo-element);
 *  - throws if any selector is still without [data-ov-theme-root].
 */
import postcss, { type AtRule, type Plugin, type Root, type Rule } from 'postcss'

export const THEME_ROOT = '[data-ov-theme-root]'
const SCOPE = `:is(${THEME_ROOT}, ${THEME_ROOT} *)`
const FALLBACK_SELECTORS = [THEME_ROOT, `${THEME_ROOT} *`, `${THEME_ROOT} ::before`, `${THEME_ROOT} ::after`]
const UTILITY_PREFIX = '.ov\\:'

/** True for the officeview stylesheet entry, with or without a Vite `?query`
 * or `#hash`, on POSIX or Windows paths. */
export function isOvEntry(id: string | undefined): boolean {
  if (!id) return false
  const path = id.replace(/[?#].*$/s, '').replace(/\\/g, '/')
  return path.endsWith('/src/styles/officeview.css')
}

/** End index of the CSS class name that starts at `start` (just after the
 * dot): identifier characters and backslash escapes. */
function classNameEnd(selector: string, start: number): number {
  let i = start
  while (i < selector.length) {
    const ch = selector[i]
    if (ch === '\\') { i += 2; continue }
    if (/[\w-]/.test(ch) || ch.charCodeAt(0) > 0x7f) { i++; continue }
    break
  }
  return i
}

function scopeSelector(selector: string): string {
  if (selector.includes(THEME_ROOT)) return selector
  if (!selector.startsWith(UTILITY_PREFIX)) return selector
  const end = classNameEnd(selector, 1)
  return selector.slice(0, end) + SCOPE + selector.slice(end)
}

function insideKeyframes(rule: Rule): boolean {
  for (let p: postcss.Node | undefined = rule.parent; p; p = p.parent) {
    if (p.type === 'atrule' && /keyframes$/i.test((p as AtRule).name)) return true
  }
  return false
}

function selectorsWithoutScope(root: Root): string[] {
  const bad: string[] = []
  root.walkRules((rule) => {
    if (insideKeyframes(rule)) return
    for (const selector of rule.selectors) if (!selector.includes(THEME_ROOT)) bad.push(selector)
  })
  return bad
}

/** Every selector (keyframe steps aside) that lacks [data-ov-theme-root]. */
export function unscopedSelectors(css: string): string[] {
  return selectorsWithoutScope(postcss.parse(css))
}

function transform(root: Root): void {
  root.walkAtRules('property', (rule) => { rule.remove() })

  // The --tw-* fallback: collect its declarations, drop the layer.
  const fallback = postcss.rule({ selectors: FALLBACK_SELECTORS })
  root.walkAtRules('layer', (layer) => {
    if (layer.params !== 'properties') return
    layer.walkRules((rule) => {
      rule.each((node) => { if (node.type === 'decl') fallback.append(node.clone()) })
    })
    layer.remove()
  })

  root.walkRules((rule) => {
    if (insideKeyframes(rule)) return
    rule.selectors = rule.selectors.map(scopeSelector)
  })

  if (fallback.nodes.length > 0) {
    const firstUtility = root.nodes.find((node) => containsUtility(node))
    if (firstUtility) root.insertBefore(firstUtility, fallback)
    else root.append(fallback)
  }

  const bad = selectorsWithoutScope(root)
  if (bad.length > 0) throw new Error(`[ov-scope] unscoped selectors in officeview.css:\n  ${bad.join('\n  ')}`)
}

function containsUtility(node: postcss.ChildNode): boolean {
  if (node.type === 'rule') return node.selectors.some((s) => s.startsWith(UTILITY_PREFIX))
  if (node.type === 'atrule') return (node.nodes ?? []).some(containsUtility)
  return false
}

/** The plugin's transform applied to a CSS string (unit tests). */
export function scopeOvCss(css: string): string {
  const root = postcss.parse(css)
  transform(root)
  return root.toString()
}

export function ovCssScope(): Plugin {
  return {
    postcssPlugin: 'ov-scope',
    OnceExit(root) {
      if (!isOvEntry(root.source?.input.file)) return
      transform(root)
    },
  }
}
ovCssScope.postcss = true as const
