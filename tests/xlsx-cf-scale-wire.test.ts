import { expect, it } from 'vitest'
import { prepareCfWireScope } from '../src/xlsx/cf/text'
import type { CfReadResult } from '../src/xlsx/view/contracts'

const stamp = { owner: 't', sheetId: '1', workbookValueGen: 1, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 0 }
const budget = { reserve(_n: number) {}, release(_n: number) {} }
const result: CfReadResult = {
  kind: 'ready', stamp, sheetId: '1', definitions: [], definitionsVersion: 1, diagnostics: [],
  states: [
    { row: 1, col: 1, properties: {}, scale: { ruleId: 7, argb: 0xffff0000, fraction: 0 } },
    { row: 2, col: 1, properties: {}, scale: { ruleId: 7, argb: 0xff00ff00, fraction: 1 } },
    { row: 3, col: 1, properties: { 'fill.pattern': 'solid', 'fill.bg': 'FF123456' }, scale: { ruleId: 7, argb: 0xff00ff00, fraction: 1 } },
  ],
}

it('wire scale states paint their fill on the effective style and the paged paint layer', async () => {
  const scope = await prepareCfWireScope(async () => ({ result, release() {} }), [{ r0: 0, r1: 10, c0: 0, c1: 2 }], { budget })
  expect(scope!.effectiveStyle(1, 1, undefined)?.fillColor).toBe('FFFF0000')
  expect(scope!.effectiveStyle(2, 1, undefined)?.fillColor).toBe('FF00FF00')
  expect(scope!.paintLayer!.styleAt(1, 1, undefined, 'screen' as never)?.fillColor).toBe('FFFF0000')
  expect(scope!.regions).toEqual(expect.arrayContaining([
    { minRow: 1, maxRow: 1, minCol: 1, maxCol: 1 }, { minRow: 2, maxRow: 2, minCol: 1, maxCol: 1 },
  ]))
  // An explicit DXF fill winning the slot is not overridden by the scale visual.
  expect(scope!.effectiveStyle(3, 1, undefined)?.fillColor?.toUpperCase()).toMatch(/123456$/)
  scope!.dispose()
})
