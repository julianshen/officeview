import { describe, expect, it } from 'vitest'
import type { CfRule } from '../src/xlsx/cf/model'
import * as dxf from '../src/xlsx/cf/dxf'
import * as composition from '../src/xlsx/cf/compose'
const compose = composition.composeCell
const yes = { kind: 'ready', value: true } as const
const unknown = { kind: 'unresolved', reason: 'cf-unsupported-rule' } as const
function rule(order: number, properties: Record<string, string | boolean | null>, writes = Object.keys(properties), stop = false): CfRule {
  return { id: order, priority: 1, documentOrder: order, ranges: [{ firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 }], stopIfTrue: stop, writes: new Set(writes), dxf: properties, definition: { kind: 'predicate', type: 'containsErrors', dialect: 'base', formulas: [] } }
}
describe('CF atomic composition', () => {
  it('unknown first rule with stop suppresses equal-priority successor', async () => {
    expect((await compose([rule(0, {}, ['fill'], true), rule(1, { 'fill.fg': 'green' }, ['fill'])], [unknown, yes])).properties).toEqual({})
  })
  it('unknown fill blocks green but permits lower bold and preserves higher blue', async () => {
    const r = [rule(0, { 'fill.fg': 'blue' }, ['fill']), rule(1, {}, ['fill']), rule(2, { 'fill.fg': 'green', 'font.bold': true }, ['fill', 'font.bold'])]
    expect((await compose(r, [yes, unknown, yes])).properties).toEqual({ 'fill.fg': 'blue', 'font.bold': true })
    expect((await compose(r.slice(1), [unknown, yes])).properties).toEqual({ 'font.bold': true })
  })
  it('unresolved dxf and future visual stubs preserve their taint', async () => {
    expect((await compose([rule(0, {}, ['ALL']), rule(1, { 'font.bold': true })], [unknown, yes])).properties).toEqual({})
    const state = await compose([rule(0, {}, ['dataBar']), rule(1, { 'font.bold': true })], [unknown, yes])
    expect(state.properties).toEqual({ 'font.bold': true }); expect(state.unresolvedWrites?.has('dataBar')).toBe(true)
  })
  it('unresolved colorScale blocks lower-priority fill while allowing font (I1)', async () => {
    const scale = rule(0, {}, ['colorScale', 'fill'])
    const cellIs = rule(1, { 'fill.fg': 'FFFF0000', 'font.bold': true }, ['fill', 'font.bold'])
    const state = await compose([scale, cellIs], [unknown, yes])
    expect(state.properties).toEqual({ 'font.bold': true })
    expect(state.unresolvedWrites?.has('fill')).toBe(true)
    expect(state.unresolvedWrites?.has('colorScale')).toBe(true)
  })
  it('bare true stop halts; no-match and false do not; earlier explicit clears win atomically', async () => {
    const r = [rule(0, {}, [], true), rule(1, { 'font.bold': true })]
    expect((await compose(r, [yes, yes])).properties).toEqual({})
    expect((await compose(r, [{ kind: 'no-match' }, yes])).properties).toEqual({ 'font.bold': true })
    expect((await compose([rule(0, { 'fill.pattern': 'none' }, ['fill']), rule(1, { 'fill.fg': 'green' }, ['fill'])], [yes, yes])).properties).toEqual({ 'fill.pattern': 'none' })
  })
  it('returned properties and write masks cannot mutate published state', async () => {
    const state = await compose([rule(0, {}, ['font']), rule(1, { 'font.bold': true })], [unknown, yes])
    expect(Object.isFrozen(state.properties)).toBe(true)
    expect((state.unresolvedWrites as Set<string>).add).toBeUndefined()
  })
  it('winning per-cell deltas override own styles without mutating the base or sharing nested edges', () => {
    expect(dxf).toHaveProperty('resolveCfCellStyle')
    const base = { numFmtId: 4, bold: true, ownFont: true as const, fillColor: 'red', borders: { top: 'thin' }, borderColors: { top: 'red' }, formatCode: 'old' }
    const before = structuredClone(base)
    const out = dxf.resolveCfCellStyle(base, { properties: { 'font.bold': false, 'fill.pattern': 'none', 'border.top.style': 'none', 'numFmt.id': 0 } })
    expect(out.bold).toBe(false); expect(out.fillColor).toBeUndefined(); expect(out.borders?.top).toBeUndefined(); expect(out.formatCode).toBeUndefined()
    expect(base).toEqual(before); expect(out.borders).not.toBe(base.borders)
  })
})

it('composition checkpoints during a large rule chain rather than after synchronous work', async () => {
  const api = composition
  expect(api).toHaveProperty('composeCellAsync')
  let reads = 0, steps = 0
  const rules = Array.from({ length: 1000 }, (_,i) => ({ ...rule(i, {}), get stopIfTrue() { reads++; return false } }))
  await expect(api.composeCellAsync(rules, rules.map(() => yes), async () => { if (++steps === 10) throw Error('cancelled') })).rejects.toThrow('cancelled')
  expect(reads).toBeLessThan(20)
})
