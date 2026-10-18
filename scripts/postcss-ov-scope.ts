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
 *  - throws if any selector could still match outside the theme root
 *    (isScopedSelector).
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

/** Top-level pieces of a complex selector: compounds and the combinators
 * between them (' ', '>', '+', '~'). Parentheses, brackets, quotes and
 * escapes are skipped over. */
function splitComplex(selector: string): { compounds: string[]; combinators: string[] } {
  const compounds: string[] = []
  const combinators: string[] = []
  let depth = 0
  let current = ''
  let pending: string | null = null
  let quote: string | null = null
  const flush = () => {
    if (current === '') return
    if (compounds.length > 0) combinators.push(pending ?? ' ')
    compounds.push(current)
    current = ''
    pending = null
  }
  for (let i = 0; i < selector.length; i++) {
    const ch = selector[i]
    if (quote) { current += ch; if (ch === '\\') current += selector[++i] ?? ''; else if (ch === quote) quote = null; continue }
    if (ch === '\\') { current += ch + (selector[++i] ?? ''); continue }
    if (ch === '"' || ch === "'") { quote = ch; current += ch; continue }
    if (ch === '(' || ch === '[') depth++
    else if (ch === ')' || ch === ']') depth--
    if (depth === 0 && /[\s>+~]/.test(ch)) {
      flush()
      if (ch !== ' ' && ch !== '\n' && ch !== '\t') pending = ch
      continue
    }
    current += ch
  }
  flush()
  return { compounds, combinators }
}

/** Splits on top-level commas. */
function splitList(list: string): string[] {
  const out: string[] = []
  let depth = 0
  let current = ''
  for (let i = 0; i < list.length; i++) {
    const ch = list[i]
    if (ch === '\\') { current += ch + (list[++i] ?? ''); continue }
    if (ch === '(' || ch === '[') depth++
    else if (ch === ')' || ch === ']') depth--
    if (depth === 0 && ch === ',') { out.push(current.trim()); current = ''; continue }
    current += ch
  }
  if (current.trim()) out.push(current.trim())
  return out
}

/** Top-level functional pseudo-classes of a compound: [name, argument]. */
function functionalPseudos(compound: string): [string, string][] {
  const out: [string, string][] = []
  let depth = 0
  for (let i = 0; i < compound.length; i++) {
    const ch = compound[i]
    if (ch === '\\') { i++; continue }
    if (ch === '[') { depth++; continue }
    if (ch === ']') { depth--; continue }
    if (depth === 0 && ch === ':') {
      const m = /^:([\w-]+)\(/.exec(compound.slice(i))
      if (!m) continue
      let level = 0
      let j = i + m[0].length - 1
      for (; j < compound.length; j++) {
        if (compound[j] === '\\') { j++; continue }
        if (compound[j] === '(') level++
        else if (compound[j] === ')' && --level === 0) break
      }
      out.push([m[1].toLowerCase(), compound.slice(i + m[0].length, j)])
      i = j
    }
  }
  return out
}

/** Whether the compound itself pins its element to the theme root or a
 * descendant: a top-level [data-ov-theme-root] attribute, or an :is()/:where()
 * group every argument of which is itself a scoped selector. */
function scopesCompound(compound: string): boolean {
  let depth = 0
  for (let i = 0; i < compound.length; i++) {
    const ch = compound[i]
    if (ch === '\\') { i++; continue }
    if (ch === '(') depth++
    else if (ch === ')') depth--
    else if (depth === 0 && compound.startsWith(THEME_ROOT, i)) return true
  }
  return functionalPseudos(compound).some(([name, arg]) => (name === 'is' || name === 'where') && splitList(arg).every(isScopedSelector))
}

/** Every negation or relational argument anywhere in the selector. */
function negatedArguments(selector: string): string[] {
  const out: string[] = []
  const re = /:(not|has)\(/gi
  let m: RegExpExecArray | null
  while ((m = re.exec(selector))) {
    let level = 1
    let j = m.index + m[0].length
    for (; j < selector.length && level > 0; j++) {
      if (selector[j] === '\\') { j++; continue }
      if (selector[j] === '(') level++
      else if (selector[j] === ')') level--
    }
    out.push(selector.slice(m.index + m[0].length, j - 1))
  }
  return out
}

/**
 * A complex selector is scoped when it can only match the theme root or its
 * descendants:
 *  - its leftmost compound pins the theme root (a top-level
 *    [data-ov-theme-root], or an :is()/:where() whose arguments are all scoped).
 *    Ancestor-prefixed forms such as `body [data-ov-theme-root]` are rejected:
 *    the library never needs host ancestors, and a leftmost-only rule keeps
 *    the gate simple and fail-closed;
 *  - no sibling combinator follows that compound (`[root] ~ div` and
 *    `[root] + *` reach host siblings); siblings further right are siblings of
 *    descendants, still inside the root;
 *  - the root marker never sits inside :not() or :has(), which would select
 *    elements outside the root (`:not([root]) button`, `:has([root])`).
 */
export function isScopedSelector(selector: string): boolean {
  if (negatedArguments(selector).some((arg) => arg.includes('[data-ov-theme-root'))) return false
  const { compounds, combinators } = splitComplex(selector.trim())
  if (compounds.length === 0 || !scopesCompound(compounds[0])) return false
  if (combinators[0] === '+' || combinators[0] === '~') return false
  return true
}

function scopeSelector(selector: string): string {
  if (isScopedSelector(selector)) return selector
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
    for (const selector of rule.selectors) if (!isScopedSelector(selector)) bad.push(selector)
  })
  return bad
}

/** Every selector (keyframe steps aside) that could match outside the theme
 * root (see isScopedSelector). */
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
