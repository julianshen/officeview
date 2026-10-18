/**
 * Task 1 (spec §6.8): the PostCSS scoping stage. Every selector the shipped
 * stylesheet emits must contain [data-ov-theme-root]; @property is dropped and
 * the --tw-* fallback becomes an unlayered, scoped rule.
 */
import { describe, expect, it } from 'vitest'
import postcss from 'postcss'
import css from '../src/styles/officeview.css?inline'
import { isOvEntry, scopeOvCss, unscopedSelectors } from '../scripts/postcss-ov-scope'

const SCOPE = ':is([data-ov-theme-root], [data-ov-theme-root] *)'

// Real tailwindcss 4.3.3 output for prefix(ov) (captured from a scratch compile).
const PROPERTIES_OUTPUT = `@layer properties;
.ov\\:border {
  border-style: var(--tw-border-style);
  border-width: 1px;
}
@property --tw-border-style {
  syntax: "*";
  inherits: false;
  initial-value: solid;
}
@property --tw-shadow {
  syntax: "*";
  inherits: false;
  initial-value: 0 0 #0000;
}
@layer properties {
  @supports ((-webkit-hyphens: none) and (not (margin-trim: inline))) or ((-moz-orient: inline) and (not (color:rgb(from red r g b)))) {
    *, ::before, ::after, ::backdrop {
      --tw-border-style: solid;
      --tw-shadow: 0 0 #0000;
    }
  }
}`

function selectorsOf(out: string): string[] {
  const list: string[] = []
  postcss.parse(out).walkRules((rule) => { list.push(...rule.selectors) })
  return list
}

describe('ovCssScope (spec §6.8)', () => {
  it('drops @property and rescopes the properties fallback', () => {
    const out = scopeOvCss(PROPERTIES_OUTPUT)
    expect(out).not.toContain('@property')
    expect(out).not.toContain('@layer')
    expect(out).not.toContain('@supports')
    const root = postcss.parse(out)
    const rules: postcss.Rule[] = []
    root.walkRules((r) => { rules.push(r) })
    // The fallback is unlayered, unconditional and placed before the utilities.
    expect(rules[0].parent?.type).toBe('root')
    expect(rules[0].selectors).toEqual([
      '[data-ov-theme-root]',
      '[data-ov-theme-root] *',
      '[data-ov-theme-root] ::before',
      '[data-ov-theme-root] ::after',
    ])
    const decls: string[] = []
    rules[0].walkDecls((d) => { decls.push(`${d.prop}: ${d.value}`) })
    expect(decls).toEqual(['--tw-border-style: solid', '--tw-shadow: 0 0 #0000'])
    expect(rules[1].selector).toBe(`.ov\\:border${SCOPE}`)
    expect(unscopedSelectors(out)).toEqual([])
  })

  it('appends the theme-root scope to utility selectors, variants and pseudo-elements', () => {
    const table: [input: string, expected: string][] = [
      ['.ov\\:flex { display: flex; }', `.ov\\:flex${SCOPE}`],
      ['@media (hover: hover) { .ov\\:hover\\:bg-hover:hover { background-color: var(--ov-color-hover); } }', `.ov\\:hover\\:bg-hover${SCOPE}:hover`],
      ['.ov\\:placeholder\\:text-muted::placeholder { color: var(--ov-color-muted); }', `.ov\\:placeholder\\:text-muted${SCOPE}::placeholder`],
      ['@media (pointer: coarse) { .ov\\:coarse\\:min-h-11 { min-height: calc(4px * 11); } }', `.ov\\:coarse\\:min-h-11${SCOPE}`],
      ['@media (prefers-reduced-motion: reduce) { .ov\\:motion-reduce\\:transition-none { transition-property: none; } }', `.ov\\:motion-reduce\\:transition-none${SCOPE}`],
      ['@media (hover: hover) { .ov\\:group-hover\\:flex:is(:where(.ov\\:group):hover *) { display: flex; } }', `.ov\\:group-hover\\:flex${SCOPE}:is(:where(.ov\\:group):hover *)`],
      ['.ov\\:bg-accent\\/50 { background-color: var(--ov-color-accent); }', `.ov\\:bg-accent\\/50${SCOPE}`],
    ]
    for (const [input, expected] of table) {
      const out = scopeOvCss(input)
      expect(selectorsOf(out), input).toEqual([expected])
      expect(unscopedSelectors(out), input).toEqual([])
    }
    // At-rule wrappers survive untouched.
    expect(scopeOvCss(table[3][0])).toContain('@media (pointer: coarse)')
    // Author rules that already carry the scope are left alone.
    const base = '[data-ov-theme-root] :is(button, input) { font: inherit; }'
    expect(selectorsOf(scopeOvCss(base))).toEqual(['[data-ov-theme-root] :is(button, input)'])
  })

  it('unscopedSelectors reports any rule without [data-ov-theme-root] (mutation)', () => {
    expect(unscopedSelectors('.x{}')).toEqual(['.x'])
    expect(unscopedSelectors(':root{}')).toEqual([':root'])
    expect(unscopedSelectors('[data-ov-theme-root], .y {}')).toEqual(['.y'])
    expect(unscopedSelectors('@media print { .z {} }')).toEqual(['.z'])
    expect(unscopedSelectors(`.ov\\:flex${SCOPE}{}`)).toEqual([])
    // keyframe steps are not selectors
    expect(unscopedSelectors('@keyframes spin { from { opacity: 0 } to { opacity: 1 } }')).toEqual([])
    // A selector the transform cannot scope fails loudly instead of shipping unscoped.
    expect(() => scopeOvCss(':where(.ov\\:space-x-2 > :not(:last-child)) { margin: 0 }')).toThrow(/unscoped/)
  })

  it('isOvEntry matches the entry with or without Vite queries and nothing else', () => {
    const table: [string | undefined, boolean][] = [
      ['/a/src/styles/officeview.css', true],
      ['/a/src/styles/officeview.css?inline', true],
      ['/a/src/styles/officeview.css?direct', true],
      ['/a/src/styles/officeview.css?inline&used', true],
      ['C:\\a\\src\\styles\\officeview.css?inline', true],
      ['/a/src/styles/officeview.css#x', true],
      ['/a/src/styles/other.css?inline', false],
      ['/a/src/styles/officeview.css.map', false],
      ['/a/src/other/officeview.css', false],
      [undefined, false],
    ]
    for (const [id, expected] of table) expect(isOvEntry(id), String(id)).toBe(expected)
  })

  it('the ?inline compiled stylesheet is scoped (pipeline placement)', () => {
    expect(css.length).toBeGreaterThan(0)
    expect(unscopedSelectors(css)).toEqual([])
    expect(css).not.toContain('@property')
    expect(css).not.toMatch(/:root|:host/)
    const scoped = selectorsOf(css).filter((s) => s.startsWith('.ov\\:') && s.endsWith(SCOPE))
    expect(scoped.length).toBeGreaterThan(0)
  })
})
