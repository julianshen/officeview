// @vitest-environment node
/**
 * Task C1: the README's "Charts" section lists exactly the families the engine draws natively: the classic
 * family registry (`CHART_GROUP_FAMILIES`, plot-area element to family) and the painted ChartEx families
 * (`CHARTEX_PAINTED`). A family added to or removed from the engine without the README fails here.
 */
import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { CHART_GROUP_FAMILIES } from '../src/charts/parse'
import { CHARTEX_PAINTED } from '../src/charts/chartex/parse'

const readme = readFileSync(join(__dirname, '..', 'README.md'), 'utf8')
/** The rows of the first table after `intro` inside the README's "## Charts" section: [first cell, second cell]. */
function table(intro: string): Array<[string, string]> {
  const section = readme.slice(readme.indexOf('\n## Charts\n'), readme.indexOf('\n## ', readme.indexOf('\n## Charts\n') + 1))
  const start = section.indexOf(intro)
  expect(start, `README Charts section: "${intro}"`).toBeGreaterThan(0)
  const rows: Array<[string, string]> = []
  let seen = false
  for (const line of section.slice(start).split('\n').slice(1)) {
    if (!line.startsWith('|')) { if (seen) break; continue }
    seen = true
    const cells = line.split('|').slice(1, -1).map(c => c.trim())
    if (cells[0] === 'Family' || /^-+$/.test(cells[0])) continue
    rows.push([cells[0], cells[1]])
  }
  return rows
}
const codes = (text: string) => [...text.matchAll(/`([A-Za-z0-9]+)`/g)].map(m => m[1])

describe('README Charts section', () => {
  it('README chart family list matches the engine family registry', () => {
    // Classic: every family once, and per family exactly the plot-area elements the registry maps to it.
    const classic = table('Classic DrawingML families')
    const families = classic.map(([family]) => codes(family)[0])
    expect(new Set(families).size).toBe(families.length)
    expect(new Set(families)).toEqual(new Set(Object.values(CHART_GROUP_FAMILIES)))
    for (const [family, elements] of classic) {
      const name = codes(family)[0]
      const registered = Object.entries(CHART_GROUP_FAMILIES).filter(([, f]) => f === name).map(([element]) => element)
      expect(new Set(codes(elements)), `README elements of ${name}`).toEqual(new Set(registered))
    }
    // ChartEx: exactly the painted families.
    const chartEx = table('Office 2016+ ChartEx families').map(([family]) => codes(family)[0])
    expect(new Set(chartEx)).toEqual(new Set(CHARTEX_PAINTED))
    expect(chartEx).toHaveLength(CHARTEX_PAINTED.size)
  })
})
