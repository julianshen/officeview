/**
 * Shipped-stylesheet audit (spec §6.8).
 *
 *   bun scripts/audit-css.ts dist/style.css
 *
 * Exits 1 if any selector lacks [data-ov-theme-root], or if the file contains
 * an @property rule, :root or :host.
 */
import { readFileSync } from 'node:fs'
import postcss from 'postcss'
import { unscopedSelectors } from './postcss-ov-scope'

const file = process.argv[2]
if (!file) {
  console.error('usage: bun scripts/audit-css.ts <file.css>')
  process.exit(2)
}
const css = readFileSync(file, 'utf8')
const problems: string[] = []
for (const selector of unscopedSelectors(css)) problems.push(`unscoped selector: ${selector}`)
const root = postcss.parse(css)
root.walkAtRules('property', (rule) => { problems.push(`@property ${rule.params}`) })
root.walkRules((rule) => {
  for (const selector of rule.selectors) if (/:root\b|:host\b/.test(selector)) problems.push(`global selector: ${selector}`)
})
if (problems.length > 0) {
  console.error(`${file}: ${problems.length} problem(s)\n  ${problems.join('\n  ')}`)
  process.exit(1)
}
let rules = 0
root.walkRules(() => { rules++ })
console.log(`${file}: ok (${rules} rules, all scoped)`)
