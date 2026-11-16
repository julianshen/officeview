// @vitest-environment node
/**
 * Stage 5, Task 35, Ruling S5-O: a resolved formula rule composes as a DXF rule, and every cross-kind outcome with a formula
 * rule on one side needs its own `formula:composition-*` capability; without it the lower contribution is withheld (its
 * writes tainted), never silently halted or painted. The Stage 2-4 compose suites run unchanged beside this file.
 */
import { describe, expect, it } from 'vitest'
import { composeCell } from '../src/xlsx/cf/compose'
import { COLOR_SCALE_WRITES, DATA_BAR_WRITES, ICON_SET_WRITES, type CfCellState, type CfOutcome, type CfRule } from '../src/xlsx/cf/model'
import { FORMULA_CAPS, GREEN, RED, parseRules, prepare, sheetOf, shown } from './helpers/cf-formula-rules'

type Kind = 'expr' | 'dxf' | 'scale' | 'bar' | 'icon'
const FILL = { 'fill.pattern': 'solid', 'fill.bg': RED } as const
const STRIKE = { 'font.strike': true } as const
function rule(id: number, kind: Kind, o: { stop?: boolean; dxf?: Readonly<Record<string, string | boolean>> } = {}): CfRule {
  const base = { id, priority: id, documentOrder: id, ranges: [{ firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 }], stopIfTrue: o.stop ?? false }
  const dxf = o.dxf ?? (kind === 'expr' || kind === 'dxf' ? FILL : undefined)
  const writes = new Set<string>(dxf ? Object.keys(dxf).map(k => (k.startsWith('fill.') ? 'fill' : k)) : [])
  const origin = { dialect: 'base' as const, formulas: ['1'] }
  switch (kind) {
    case 'expr': return { ...base, writes, dxf: dxf!, definition: { ...origin, kind: 'expression', formula: {} as never, sourceType: 'expression' } }
    case 'dxf': return { ...base, writes, dxf: dxf!, definition: { ...origin, kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 0 }] } }
    case 'scale': return { ...base, writes: COLOR_SCALE_WRITES, definition: { ...origin, kind: 'colorScale', points: [{ bound: { type: 'min' }, argb: 0xfff8696b }, { bound: { type: 'max' }, argb: 0xff63be7b }] } }
    case 'bar': return { ...base, writes: DATA_BAR_WRITES, definition: { ...origin, kind: 'dataBar' } as never }
    case 'icon': return { ...base, writes: ICON_SET_WRITES, definition: { ...origin, kind: 'iconSet' } as never }
  }
}
const ready = (kind: Kind): CfOutcome<unknown> => kind === 'scale' ? { kind: 'ready', value: { argb: 0xff63be7b, fraction: 1 } }
  : kind === 'bar' ? { kind: 'ready', value: { axisFraction: 0, startFraction: 0, endFraction: 0.5, valueFraction: 0.5, fillColor: 0xff638ec6, isGradient: false, isNegative: false, showValue: true, direction: 'leftToRight' } }
  : kind === 'icon' ? { kind: 'ready', value: { set: '3Arrows', iconId: 2, bucket: 2, buckets: 3, showValue: true } }
  : { kind: 'ready', value: true }
/** The Stage 2-4 composition capabilities the lower visuals compose under (production has them verified). */
const STAGE = ['scale:composition-matching-stopIfTrue-ahead', 'scale:composition-dxf-p1-scale-p2', 'scale:composition-scale-p1-dxf-p2',
  'bar:composition-dxf-fill', 'bar:composition-matching-stopIfTrue-ahead', 'icon:composition-dxf-fill', 'icon:composition-matching-stopIfTrue-ahead', 'scale:composition-font-only-dxf']
const compose = (rules: CfRule[], outcomes: CfOutcome<unknown>[], caps: readonly string[]): CfCellState => composeCell(rules, outcomes, { hasEvidence: cap => caps.includes(cap) })
const view = (s: CfCellState) => ({
  properties: { ...s.properties }, tainted: [...(s.unresolvedWrites ?? [])].sort(),
  visual: Object.keys(s.visual ?? {}).sort(),
})

describe('formula composition across kinds (S5-O)', () => {
  describe('an expression p1 stopIfTrue true halts the lower kind only with formula:composition-stop-<kind>', () => {
    for (const lower of ['dxf', 'scale', 'bar', 'icon'] as const) {
      it(`ahead of a ${lower}`, () => {
        const cap = `formula:composition-stop-${lower}`
        const rules = [rule(1, 'expr', { stop: true, dxf: STRIKE }), rule(2, lower, lower === 'dxf' ? { dxf: { 'fill.pattern': 'solid', 'fill.bg': GREEN } } : {})]
        const outcomes = [ready('expr'), ready(lower)]
        const halted = view(compose(rules, outcomes, [...STAGE, cap]))
        expect(halted).toEqual({ properties: STRIKE, tainted: [], visual: [] })
        const withheld = view(compose(rules, outcomes, STAGE))
        expect(withheld.properties).toEqual(STRIKE) // the expression's own DXF still applies
        expect(withheld.visual).toEqual([]) // the lower contribution never paints
        expect(withheld.tainted).toEqual([...rules[1]!.writes].sort()) // ... and is withheld, not silently halted
        // Another kind's capability never stands in.
        const other = (['dxf', 'scale', 'bar', 'icon'] as const).filter(k => k !== lower).map(k => `formula:composition-stop-${k}`)
        expect(view(compose(rules, outcomes, [...STAGE, ...other])).tainted).toEqual([...rules[1]!.writes].sort())
      })
    }
    it('a lower rule that does not match is not withheld; a literal stopIfTrue ahead is unchanged (Stage 1)', () => {
      const rules = [rule(1, 'expr', { stop: true, dxf: STRIKE }), rule(2, 'dxf', { dxf: { 'fill.pattern': 'solid', 'fill.bg': GREEN } })]
      expect(view(compose(rules, [ready('expr'), { kind: 'no-match' }], STAGE))).toEqual({ properties: STRIKE, tainted: [], visual: [] })
      const literal = [rule(1, 'dxf', { stop: true, dxf: STRIKE }), rule(2, 'dxf', { dxf: { 'fill.pattern': 'solid', 'fill.bg': GREEN } })]
      expect(view(compose(literal, [ready('dxf'), ready('dxf')], STAGE))).toEqual({ properties: STRIKE, tainted: [], visual: [] })
    })
    it('an expression that is false does not stop: the lower rule composes as usual', () => {
      const rules = [rule(1, 'expr', { stop: true, dxf: STRIKE }), rule(2, 'dxf', { dxf: { 'fill.pattern': 'solid', 'fill.bg': GREEN } })]
      expect(view(compose(rules, [{ kind: 'no-match' }, ready('dxf')], STAGE)).properties).toEqual({ 'fill.pattern': 'solid', 'fill.bg': GREEN })
    })
  })

  describe('an expression fill against a scale fill follows formula:composition-expression-vs-scale', () => {
    const cap = 'formula:composition-expression-vs-scale'
    it('expression first: the expression fill wins only with the capability; otherwise the fill is withheld', () => {
      const rules = [rule(1, 'expr'), rule(2, 'scale')]
      const outcomes = [ready('expr'), ready('scale')]
      expect(view(compose(rules, outcomes, [...STAGE, cap]))).toEqual({ properties: FILL, tainted: [], visual: [] })
      const withheld = view(compose(rules, outcomes, STAGE))
      expect(withheld).toEqual({ properties: {}, tainted: ['colorScale', 'fill'], visual: [] })
    })
    it('scale first: the scale wins (priority-wins) only with the capability', () => {
      const rules = [rule(1, 'scale'), rule(2, 'expr')]
      const outcomes = [ready('scale'), ready('expr')]
      expect(view(compose(rules, outcomes, [...STAGE, cap]))).toEqual({ properties: {}, tainted: [], visual: ['colorScale'] })
      expect(view(compose(rules, outcomes, STAGE))).toEqual({ properties: {}, tainted: ['colorScale', 'fill'], visual: [] })
    })
    it('a literal DXF against a scale is unchanged (Stage 2): it never needs the formula capability', () => {
      const rules = [rule(1, 'dxf'), rule(2, 'scale')]
      expect(view(compose(rules, [ready('dxf'), ready('scale')], STAGE))).toEqual({ properties: FILL, tainted: [], visual: [] })
    })
  })

  it('an unresolved expression p1 taints exactly its DXF writes; a lower orthogonal write still paints', () => {
    const rules = [rule(1, 'expr'), rule(2, 'dxf', { dxf: { 'font.bold': true, 'fill.pattern': 'solid', 'fill.bg': GREEN } })]
    const out = view(compose(rules, [{ kind: 'unresolved', reason: 'cf-formula-unsupported' }, ready('dxf')], STAGE))
    expect(out.tainted).toEqual([...rules[0]!.writes].sort())
    expect(out.properties).toEqual({ 'font.bold': true })
  })

  describe('two overlapping expressions write first-writer-wins only with formula:composition-overlapping and the Stage 1 guard', () => {
    const xml = `<conditionalFormatting sqref="P2:P3"><cfRule type="expression" priority="1" dxfId="0"><formula>$A2&gt;0</formula></cfRule>`
      + `<cfRule type="expression" priority="2" dxfId="1"><formula>$A2&gt;-1</formula></cfRule></conditionalFormatting>`
    const STAGE1 = ['priority:fill-winner', 'priority:orthogonal-font-bold']
    const run = async (caps: readonly string[]) => {
      const out = await prepare(sheetOf(parseRules(xml, FORMULA_CAPS), { A2: 1, A3: 0, P2: 1, P3: 1 }), ['P2', 'P3'], caps)
      return ['P2', 'P3'].map(r => { const s = shown(out.get(r)); return s.withheld || out.get(r)?.kind === 'unresolved' ? 'withheld' : s.fill })
    }
    it('both true: red (first) with every capability; withheld without either', async () => {
      expect(await run([...FORMULA_CAPS, ...STAGE1, 'formula:composition-overlapping'])).toEqual([RED, GREEN])
      expect(await run([...FORMULA_CAPS, ...STAGE1])).toEqual(['withheld', GREEN])
      expect(await run([...FORMULA_CAPS, 'formula:composition-overlapping'])).toEqual(['withheld', GREEN])
    })
  })
})
