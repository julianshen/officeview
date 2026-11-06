import { expect, it, vi } from 'vitest'
import { parseXml } from '../src/core/xml'
import { ensureCf, ensureCfParser, hasConditionalFormatting, CF_PRESENCE_EXTENSION_URI } from '../src/xlsx/cfLoader'
import { CF_X14_EXTENSION_URI } from '../src/xlsx/cf/parse'
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
  for (const rules of [EAGER_FORBIDDEN, PAGED_WORKER_FORBIDDEN]) for (const file of ['model', 'limits', 'evidence', 'parse', 'capture', 'dxf', 'predicates']) {
    expect(rules.some(rule => rule.test(`/repo/src/xlsx/cf/${file}.ts`))).toBe(true)
  }
})
it('shares one ext URI between the eager presence check and the lazy parser', () => {
  expect(CF_PRESENCE_EXTENSION_URI).toBe(CF_X14_EXTENSION_URI)
})
it('loads the rule parser only when CF content is present', async () => {
  const load = vi.fn(async () => ({}))
  expect(await ensureCfParser(false, load)).toBeUndefined()
  expect(load).not.toHaveBeenCalled()
  expect(await ensureCfParser(true, load)).toEqual({})
  expect(load).toHaveBeenCalledTimes(1)
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
  const plain = parseXml(`<worksheet xmlns="${NS}"><sheetData/></worksheet>`)
  expect(hasConditionalFormatting(plain)).toBe(false)
  const base = parseXml(`<worksheet xmlns="${NS}"><conditionalFormatting sqref="A1"/></worksheet>`)
  expect(hasConditionalFormatting(base)).toBe(true)
  const x14 = parseXml(
    `<worksheet xmlns="${NS}"><extLst><ext uri="${CF_PRESENCE_EXTENSION_URI}"><conditionalFormattings xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"/></ext></extLst></worksheet>`,
  )
  expect(hasConditionalFormatting(x14)).toBe(true)
  expect(hasConditionalFormatting(undefined)).toBe(false)
})
