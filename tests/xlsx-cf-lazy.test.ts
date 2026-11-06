import { expect, it, vi } from 'vitest'
import { ensureCf } from '../src/xlsx/cfLoader'
import { EAGER_FORBIDDEN, PAGED_WORKER_FORBIDDEN } from './helpers/eager-graph'
it('no CF workbook imports no CF implementation or allocates a CF cache', async () => {
  const load = vi.fn(async () => ({}))
  const sheets = [{ name: 'plain', conditionalFormatting: undefined }]
  expect(await ensureCf(sheets, load)).toBeUndefined()
  expect(load).not.toHaveBeenCalled(); expect(Object.keys(sheets[0])).toEqual(['name', 'conditionalFormatting'])
  expect(await ensureCf([{ conditionalFormatting: { rules: [], diagnostics: [] } }], load)).toEqual({})
  expect(load).toHaveBeenCalledTimes(1)
})
it('excludes every CF implementation from both ordinary static closures', () => {
  for (const rules of [EAGER_FORBIDDEN, PAGED_WORKER_FORBIDDEN]) for (const file of ['model', 'limits', 'evidence', 'parse', 'predicates']) {
    expect(rules.some(rule => rule.test(`/repo/src/xlsx/cf/${file}.ts`))).toBe(true)
  }
})
