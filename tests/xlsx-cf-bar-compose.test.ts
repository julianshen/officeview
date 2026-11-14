import { describe, expect, it } from 'vitest'
import type { CfBarDescriptor, CfDefinition, CfOutcome, CfRule } from '../src/xlsx/cf/model'
import { composeCell, visualSlotOf } from '../src/xlsx/cf/compose'
import { withExactCfEvidence } from '../src/xlsx/cf/evidence'

// Slot composition for data bars (Rulings S3-I, S3-J). Bar outcomes here are stubbed descriptors: composition decides
// only whether a resolved bar may paint, never its geometry.

const R = [{ firstRow: 0, lastRow: 10, firstCol: 0, lastCol: 10 }]
type BarDef = Extract<CfDefinition, { kind: 'dataBar' }>
const BAR_DEF: BarDef = {
  kind: 'dataBar', dialect: 'base', formulas: [], bounds: [{ type: 'min' }, { type: 'max' }], fillArgb: 0xff0000ff,
  minLength: 0, maxLength: 100, gradient: false, border: false, negativeBarColorSameAsPositive: false,
  negativeBarBorderColorSameAsPositive: true, axisPosition: 'automatic', direction: 'leftToRight', showValue: true, sheetRightToLeft: false,
}

function barRule(id: number, priority: number, stop = false): CfRule {
  return { id, priority, documentOrder: id, ranges: R, stopIfTrue: stop, writes: new Set(['dataBar']), definition: BAR_DEF }
}
function dxfRule(id: number, priority: number, dxf: Record<string, string | boolean>, stop = false): CfRule {
  return {
    id, priority, documentOrder: id, ranges: R, stopIfTrue: stop, writes: new Set(Object.keys(dxf).map(k => k.startsWith('fill.') ? 'fill' : k)), dxf,
    definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 0 }], dialect: 'base', formulas: [] },
  }
}
function scaleRule(id: number, priority: number): CfRule {
  return {
    id, priority, documentOrder: id, ranges: R, stopIfTrue: false, writes: new Set(['colorScale', 'fill']),
    definition: { kind: 'colorScale', dialect: 'base', formulas: [], points: [{ bound: { type: 'min' }, argb: 0xfff8696b }, { bound: { type: 'max' }, argb: 0xff63be7b }] },
  }
}
const unresolvedStub = (id: number, priority: number, stop: boolean): CfRule => ({
  id, priority, documentOrder: id, ranges: R, stopIfTrue: stop, writes: new Set(['fill']),
  definition: { kind: 'unresolved', type: 'expression', reason: 'cf-unsupported-rule' },
})

const descriptor: Omit<CfBarDescriptor, 'ruleId'> = Object.freeze({
  axisFraction: 0, startFraction: 0, endFraction: 0.5, valueFraction: 0.5, fillColor: 0xff0000ff, isGradient: false, isNegative: false, showValue: true, direction: 'leftToRight',
})
const barOut: CfOutcome<unknown> = { kind: 'ready', value: descriptor }
const yes: CfOutcome<unknown> = { kind: 'ready', value: true }
const no: CfOutcome<unknown> = { kind: 'no-match' }
const unknown: CfOutcome<unknown> = { kind: 'unresolved', reason: 'cf-unsupported-rule' }
const scaleOut: CfOutcome<unknown> = { kind: 'ready', value: { argb: 0xff63be7b, fraction: 1 } }
const RED_FILL = { 'fill.pattern': 'solid', 'fill.bg': 'FFFFC7CE' }

describe('CF Stage 3 Task 20 — bar slot composition (S3-I, S3-J)', () => {
  it('visualSlotOf names the visual slot of a scale, a bar and a DXF rule', () => {
    expect(visualSlotOf(scaleRule(1, 1))).toBe('colorScale')
    expect(visualSlotOf(barRule(1, 1))).toBe('dataBar')
    expect(visualSlotOf(dxfRule(1, 1, RED_FILL))).toBeUndefined()
    // An unresolved bar stub keeps its slot through the canonical write mask.
    expect(visualSlotOf({ ...unresolvedStub(1, 1, false), writes: new Set(['dataBar']) })).toBe('dataBar')
  })

  it('a lone resolved bar paints its descriptor with the rule id', () => {
    withExactCfEvidence([], () => {
      const st = composeCell([barRule(7, 1)], [barOut])
      expect(st.visual?.dataBar).toEqual({ ruleId: 7, ...descriptor })
      expect(st.unresolvedWrites).toBeUndefined()
      expect(st.properties).toEqual({})
    })
  })

  it('two matching bars → no bar, with unresolvedWrites containing dataBar', () => {
    withExactCfEvidence([], () => {
      const st = composeCell([barRule(1, 1), barRule(2, 2)], [barOut, barOut])
      expect(st.visual?.dataBar).toBeUndefined()
      expect(st.unresolvedWrites?.has('dataBar')).toBe(true)
    })
    // Ruling S3-I-OVERLAP: with bar:composition-overlapping-bars, the first matching bar in priority order paints (Excel
    // paints stage3-bar-composition E2:E4 with the first bar's FF0000FF, never the second's FFFFA500); later bars are
    // suppressed, not unresolved.
    withExactCfEvidence(['bar:composition-overlapping-bars'], () => {
      const second = { ...descriptor, fillColor: 0xffffa500 }
      for (const rules of [[barRule(1, 1), barRule(2, 2), barRule(3, 3)], [barRule(1, 1), barRule(2, 2)]]) {
        const st = composeCell(rules, rules.map((_, i) => i === 0 ? barOut : { kind: 'ready', value: second }))
        expect(st.visual?.dataBar).toEqual({ ruleId: 1, ...descriptor })
        expect(st.unresolvedWrites).toBeUndefined()
      }
      // The first MATCHING bar: a non-matching higher bar does not count.
      expect(composeCell([barRule(1, 1), barRule(2, 2), barRule(3, 3)], [no, barOut, { kind: 'ready', value: second }]).visual?.dataBar?.ruleId).toBe(2)
      // An unresolved higher bar still taints the slot.
      expect(composeCell([barRule(1, 1), barRule(2, 2), barRule(3, 3)], [unknown, barOut, barOut]).visual?.dataBar).toBeUndefined()
    })
    // One matching and one non-matching bar: the matching one paints.
    withExactCfEvidence([], () => {
      expect(composeCell([barRule(1, 1), barRule(2, 2)], [no, barOut]).visual?.dataBar?.ruleId).toBe(2)
    })
    // An unresolved higher bar taints the slot for the lower one.
    withExactCfEvidence([], () => {
      const st = composeCell([barRule(1, 1), barRule(2, 2)], [unknown, barOut])
      expect(st.visual?.dataBar).toBeUndefined()
      expect(st.unresolvedWrites?.has('dataBar')).toBe(true)
    })
  })

  it('bar + scale → the scale fill paints, and the bar only with bar:composition-with-scale', () => {
    for (const order of ['bar-first', 'scale-first'] as const) {
      const rules = order === 'bar-first' ? [barRule(1, 1), scaleRule(2, 2)] : [scaleRule(1, 1), barRule(2, 2)]
      const outs = order === 'bar-first' ? [barOut, scaleOut] : [scaleOut, barOut]
      withExactCfEvidence([], () => {
        const st = composeCell(rules, outs)
        expect(st.visual?.colorScale?.argb, order).toBe(0xff63be7b)
        expect(st.visual?.dataBar, order).toBeUndefined()
        expect(st.unresolvedWrites?.has('dataBar'), order).toBe(true)
      })
      withExactCfEvidence(['bar:composition-with-scale'], () => {
        const st = composeCell(rules, outs)
        expect(st.visual?.colorScale?.argb, order).toBe(0xff63be7b)
        expect(st.visual?.dataBar?.endFraction, order).toBe(0.5)
      })
    }
  })

  it('bar + DXF fill (both orders) → the DXF fill paints, and the bar only with bar:composition-dxf-fill', () => {
    for (const order of ['bar-first', 'dxf-first'] as const) {
      const rules = order === 'bar-first' ? [barRule(1, 1), dxfRule(2, 2, RED_FILL)] : [dxfRule(1, 1, RED_FILL), barRule(2, 2)]
      const outs = order === 'bar-first' ? [barOut, yes] : [yes, barOut]
      withExactCfEvidence([], () => {
        const st = composeCell(rules, outs)
        expect(st.properties['fill.bg'], order).toBe('FFFFC7CE')
        expect(st.visual?.dataBar, order).toBeUndefined()
      })
      withExactCfEvidence(['bar:composition-dxf-fill'], () => {
        const st = composeCell(rules, outs)
        expect(st.properties['fill.bg'], order).toBe('FFFFC7CE')
        expect(st.visual?.dataBar?.ruleId, order).toBe(order === 'bar-first' ? 1 : 2)
      })
    }
    // A non-matching DXF fill does not gate the bar; a font-only DXF never does.
    withExactCfEvidence([], () => {
      expect(composeCell([dxfRule(1, 1, RED_FILL), barRule(2, 2)], [no, barOut]).visual?.dataBar).toBeDefined()
      const font = composeCell([dxfRule(1, 1, { 'font.bold': true }), barRule(2, 2)], [yes, barOut])
      expect(font.properties['font.bold']).toBe(true)
      expect(font.visual?.dataBar).toBeDefined()
    })
  })

  it('a matching p1 DXF stopIfTrue → no bar', () => {
    for (const evidence of [[], ['bar:composition-matching-stopIfTrue-ahead', 'bar:composition-dxf-fill']]) {
      withExactCfEvidence(evidence, () => {
        const st = composeCell([dxfRule(1, 1, { 'font.bold': true }, true), barRule(2, 2)], [yes, barOut])
        expect(st.visual?.dataBar).toBeUndefined()
        expect(st.properties['font.bold']).toBe(true)
      })
    }
    // A non-matching stopIfTrue DXF ahead does not halt the bar.
    withExactCfEvidence([], () => {
      expect(composeCell([dxfRule(1, 1, { 'font.bold': true }, true), barRule(2, 2)], [no, barOut]).visual?.dataBar).toBeDefined()
    })
  })

  it('a bar with stopIfTrue="1" → tainted contains ALL, and the p2 DXF fill and font are withheld until bar:composition-bar-stopIfTrue', () => {
    const rules = [barRule(1, 1, true), dxfRule(2, 2, { ...RED_FILL, 'font.bold': true })]
    for (const barOutcome of [barOut, no, unknown]) {
      withExactCfEvidence([], () => {
        const st = composeCell(rules, [barOutcome, yes])
        expect(st.unresolvedWrites?.has('ALL')).toBe(true)
        expect(st.properties).toEqual({})
        expect(st.visual?.dataBar).toBeUndefined()
      })
    }
    // Measured hypothesis (OD-1 analogue; Excel's capture paints the later fill): the flag is ignored on a bar.
    withExactCfEvidence(['bar:composition-bar-stopIfTrue', 'bar:composition-dxf-fill'], () => {
      const st = composeCell(rules, [barOut, yes])
      expect(st.unresolvedWrites).toBeUndefined()
      expect(st.properties['fill.bg']).toBe('FFFFC7CE')
      expect(st.properties['font.bold']).toBe(true)
      expect(st.visual?.dataBar?.ruleId).toBe(1)
    })
  })

  it('an unresolved p1 stopIfTrue stub → no bar', () => {
    withExactCfEvidence(['bar:composition-unresolved-stub', 'bar:composition-dxf-fill'], () => {
      const st = composeCell([unresolvedStub(1, 1, true), barRule(2, 2)], [unknown, barOut])
      expect(st.visual?.dataBar).toBeUndefined()
      expect(st.unresolvedWrites?.has('ALL')).toBe(true)
    })
  })
})
