/**
 * Task 8 (plan Global Constraints, "Tokens"): the README's Theming section
 * documents the host-facing colour tokens. Its table must list exactly
 * COLOR_TOKENS, each with the THEME_DEFAULTS light and dark values, so the
 * documentation cannot drift from the shipped tokens.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { COLOR_TOKENS, THEME_DEFAULTS, type ColorToken } from '../src/styles/tokens'

const README = readFileSync(join(__dirname, '..', 'README.md'), 'utf8')

/** The `### Theming` section, up to the next heading of level 1–3. */
function themingSection(): string {
  const lines = README.split('\n')
  const start = lines.findIndex(line => /^### Theming\s*$/.test(line))
  if (start < 0) return ''
  const rest = lines.slice(start + 1)
  const end = rest.findIndex(line => /^#{1,3} /.test(line))
  return (end < 0 ? rest : rest.slice(0, end)).join('\n')
}

/** Rows shaped `| \`--ov-color-<name>\` | <light> | <dark> | …`. */
function tokenRows(section: string): { name: string; light: string; dark: string }[] {
  const rows: { name: string; light: string; dark: string }[] = []
  for (const line of section.split('\n')) {
    const match = /^\|\s*`--ov-color-([a-z-]+)`\s*\|\s*`?([^|`]+?)`?\s*\|\s*`?([^|`]+?)`?\s*\|/.exec(line)
    if (match) rows.push({ name: match[1], light: match[2], dark: match[3] })
  }
  return rows
}

describe('README theming documentation', () => {
  it('README token table lists exactly COLOR_TOKENS with THEME_DEFAULTS', () => {
    const section = themingSection()
    expect(section, 'README has a ### Theming section').not.toBe('')
    const rows = tokenRows(section)
    const names = rows.map(row => row.name)
    // Exactly once each: no missing, extra or duplicated token rows.
    expect([...names].sort()).toEqual([...COLOR_TOKENS].sort())
    for (const row of rows) {
      const name = row.name as ColorToken
      expect(row.light, `${name} light`).toBe(THEME_DEFAULTS.light[name])
      expect(row.dark, `${name} dark`).toBe(THEME_DEFAULTS.dark[name])
    }
    // No --ov-color-* rows outside the Theming section's table.
    const everywhere = tokenRows(README).map(row => row.name)
    expect(everywhere.length).toBe(COLOR_TOKENS.length)
  })
})
