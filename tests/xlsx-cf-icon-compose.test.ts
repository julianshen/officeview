import { afterEach, describe, expect, it } from 'vitest'
import type { CfBarDescriptor, CfDefinition, CfIconDescriptor, CfOutcome, CfRule } from '../src/xlsx/cf/model'
import { composeCell, visualSlotOf } from '../src/xlsx/cf/compose'
import { injectCfIconSelections, withExactCfEvidence } from '../src/xlsx/cf/evidence'

// Slot composition for icon sets (Rulings S4-I, S4-J). Icon outcomes are stubbed descriptors: composition decides only
// whether a resolved icon may paint. Every expectation is a hypothesis under injected evidence, never an Excel fact.

const R = [{ firstRow: 0, lastRow: 10, firstCol: 0, lastCol: 10 }]
type IconDef = Extract<CfDefinition, { kind: 'iconSet' }>
const ICON_DEF: IconDef = {
  kind: 'iconSet', dialect: 'base', formulas: [], set: '3Arrows', reverse: false, showValue: true, percentAttr: true, sheetRightToLeft: false,
  thresholds: [{ bound: { type: 'percent', value: 0 }, gte: true }, { bound: { type: 'percent', value: 33 }, gte: true }, { bound: { type: 'percent', value: 67 }, gte: true }],
}
type BarDef = Extract<CfDefinition, { kind: 'dataBar' }>
const BAR_DEF: BarDef = {
  kind: 'dataBar', dialect: 'base', formulas: [], bounds: [{ type: 'min' }, { type: 'max' }], fillArgb: 0xff0000ff,
  minLength: 0, maxLength: 100, gradient: false, border: false, negativeBarColorSameAsPositive: false,
  negativeBarBorderColorSameAsPositive: true, axisPosition: 'automatic', direction: 'leftToRight', showValue: true, sheetRightToLeft: false,
}
const iconRule = (id: number, priority: number, stop = false): CfRule => ({ id, priority, documentOrder: id, ranges: R, stopIfTrue: stop, writes: new Set(['iconSet']), definition: ICON_DEF })
const barRule = (id: number, priority: number): CfRule => ({ id, priority, documentOrder: id, ranges: R, stopIfTrue: false, writes: new Set(['dataBar']), definition: BAR_DEF })
const dxfRule = (id: number, priority: number, dxf: Record<string, string | boolean>, stop = false): CfRule => ({
  id, priority, documentOrder: id, ranges: R, stopIfTrue: stop, writes: new Set(Object.keys(dxf).map(k => k.startsWith('fill.') ? 'fill' : k)), dxf,
  definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 0 }], dialect: 'base', formulas: [] },
})
const scaleRule = (id: number, priority: number): CfRule => ({
  id, priority, documentOrder: id, ranges: R, stopIfTrue: false, writes: new Set(['colorScale', 'fill']),
  definition: { kind: 'colorScale', dialect: 'base', formulas: [], points: [{ bound: { type: 'min' }, argb: 0xfff8696b }, { bound: { type: 'max' }, argb: 0xff63be7b }] },
})
const unresolvedStub = (id: number, priority: number, stop: boolean, writes = ['fill']): CfRule => ({
  id, priority, documentOrder: id, ranges: R, stopIfTrue: stop, writes: new Set(writes),
  definition: { kind: 'unresolved', type: 'expression', reason: 'cf-unsupported-rule' },
})

const icon: Omit<CfIconDescriptor, 'ruleId'> = Object.freeze({ set: '3Arrows', iconId: 2, bucket: 2, buckets: 3, showValue: true })
const second: Omit<CfIconDescriptor, 'ruleId'> = Object.freeze({ set: '3Flags', iconId: 0, bucket: 0, buckets: 3, showValue: true })
const noIcons: Omit<CfIconDescriptor, 'ruleId'> = Object.freeze({ set: 'NoIcons', iconId: 0, bucket: 1, buckets: 3, showValue: true })
const bar: Omit<CfBarDescriptor, 'ruleId'> = Object.freeze({
  axisFraction: 0, startFraction: 0, endFraction: 0.5, valueFraction: 0.5, fillColor: 0xff0000ff, isGradient: false, isNegative: false, showValue: true, direction: 'leftToRight',
})
const ready = (value: unknown): CfOutcome<unknown> => ({ kind: 'ready', value })
const iconOut = ready(icon), barOut = ready(bar), yes = ready(true), scaleOut = ready({ argb: 0xff63be7b, fraction: 1 })
const no: CfOutcome<unknown> = { kind: 'no-match' }
const unknown: CfOutcome<unknown> = { kind: 'unresolved', reason: 'cf-unsupported-rule' }
const RED_FILL = { 'fill.pattern': 'solid', 'fill.bg': 'FFFFC7CE' }
afterEach(() => { injectCfIconSelections(null) })

describe('CF Stage 4 Task 26 — icon slot composition (S4-I, S4-J)', () => {
  it('visualSlotOf returns iconSet for an icon rule and an icon stub', () => {
    expect(visualSlotOf(iconRule(1, 1))).toBe('iconSet')
    expect(visualSlotOf({ ...unresolvedStub(1, 1, false), writes: new Set(['iconSet']) })).toBe('iconSet')
    expect(visualSlotOf(barRule(1, 1))).toBe('dataBar')
    expect(visualSlotOf(scaleRule(1, 1))).toBe('colorScale')
    expect(visualSlotOf(dxfRule(1, 1, RED_FILL))).toBeUndefined()
  })

  it('a lone resolved icon paints its descriptor with the rule id', () => {
    withExactCfEvidence([], () => {
      const st = composeCell([iconRule(7, 1)], [iconOut])
      expect(st.visual?.iconSet).toEqual({ ruleId: 7, ...icon })
      expect(st.unresolvedWrites).toBeUndefined()
      expect(st.properties).toEqual({})
    })
  })

  it('two matching icons → no icon with iconSet tainted; with overlapping evidence the first in priority order wins', () => {
    withExactCfEvidence([], () => {
      const st = composeCell([iconRule(1, 1), iconRule(2, 2)], [iconOut, ready(second)])
      expect(st.visual?.iconSet).toBeUndefined()
      expect(st.unresolvedWrites?.has('iconSet')).toBe(true)
      // One matching and one non-matching icon: the matching one paints.
      expect(composeCell([iconRule(1, 1), iconRule(2, 2)], [no, ready(second)]).visual?.iconSet?.ruleId).toBe(2)
      // An unresolved higher icon taints the slot for the lower one.
      const tainted = composeCell([iconRule(1, 1), iconRule(2, 2)], [unknown, iconOut])
      expect(tainted.visual?.iconSet).toBeUndefined()
      expect(tainted.unresolvedWrites?.has('iconSet')).toBe(true)
    })
    withExactCfEvidence(['icon:composition-overlapping-icons'], () => {
      for (const rules of [[iconRule(1, 1), iconRule(2, 2)], [iconRule(1, 1), iconRule(2, 2), iconRule(3, 3)]]) {
        const st = composeCell(rules, rules.map((_, i) => i === 0 ? iconOut : ready(second)))
        expect(st.visual?.iconSet).toEqual({ ruleId: 1, ...icon })
        expect(st.unresolvedWrites).toBeUndefined()
      }
      expect(composeCell([iconRule(1, 1), iconRule(2, 2), iconRule(3, 3)], [no, ready(second), iconOut]).visual?.iconSet?.ruleId).toBe(2)
    })
  })

  it('icon + bar → the bar keeps its Stage 3 outcome; the icon paints only with icon:composition-with-bar', () => {
    for (const order of ['icon-first', 'bar-first'] as const) {
      const rules = order === 'icon-first' ? [iconRule(1, 1), barRule(2, 2)] : [barRule(1, 1), iconRule(2, 2)]
      const outs = order === 'icon-first' ? [iconOut, barOut] : [barOut, iconOut]
      withExactCfEvidence([], () => {
        const st = composeCell(rules, outs)
        expect(st.visual?.dataBar?.endFraction, order).toBe(0.5)
        expect(st.visual?.iconSet, order).toBeUndefined()
        expect(st.unresolvedWrites?.has('iconSet'), order).toBe(true)
        expect(st.unresolvedWrites?.has('dataBar') ?? false, order).toBe(false)
      })
      withExactCfEvidence(['icon:composition-with-bar'], () => {
        const st = composeCell(rules, outs)
        expect(st.visual?.dataBar?.endFraction, order).toBe(0.5)
        expect(st.visual?.iconSet?.iconId, order).toBe(2)
        expect(st.unresolvedWrites, order).toBeUndefined()
      })
    }
    // An earlier unresolved bar is a possible bar: the icon withholds without evidence.
    withExactCfEvidence([], () => {
      expect(composeCell([{ ...unresolvedStub(1, 1, false), writes: new Set(['dataBar']) }, iconRule(2, 2)], [unknown, iconOut]).visual?.iconSet).toBeUndefined()
    })
  })

  it('bar showValue=1 + icon showValue=0 → the pair is kept and text suppression is unresolved until icon:composition-bar-showvalue', () => {
    const hidden = ready({ ...icon, showValue: false })
    for (const [b, i] of [[barOut, hidden], [ready({ ...bar, showValue: false }), iconOut]] as const) {
      withExactCfEvidence(['icon:composition-with-bar'], () => {
        const st = composeCell([barRule(1, 1), iconRule(2, 2)], [b, i])
        expect(st.visual?.dataBar).toBeDefined()
        expect(st.visual?.iconSet).toBeDefined()
        expect(st.unresolvedWrites?.has('showValue')).toBe(true)
        expect(st.unresolvedWrites?.has('iconSet')).toBe(false)
        expect(st.unresolvedWrites?.has('dataBar')).toBe(false)
      })
      withExactCfEvidence(['icon:composition-with-bar', 'icon:composition-bar-showvalue'], () => {
        expect(composeCell([barRule(1, 1), iconRule(2, 2)], [b, i]).unresolvedWrites).toBeUndefined()
      })
    }
    // Agreeing showValue never marks anything.
    withExactCfEvidence(['icon:composition-with-bar'], () => {
      expect(composeCell([barRule(1, 1), iconRule(2, 2)], [barOut, iconOut]).unresolvedWrites).toBeUndefined()
    })
  })

  it('icon + scale / DXF fill → the fill paints, and the icon only with its capability', () => {
    for (const order of ['icon-first', 'fill-first'] as const) {
      for (const [fill, out, cap] of [[scaleRule, scaleOut, 'icon:composition-with-scale'], [(id: number, p: number) => dxfRule(id, p, RED_FILL), yes, 'icon:composition-dxf-fill']] as const) {
        const rules = order === 'icon-first' ? [iconRule(1, 1), fill(2, 2)] : [fill(1, 1), iconRule(2, 2)]
        const outs = order === 'icon-first' ? [iconOut, out] : [out, iconOut]
        const painted = (st: ReturnType<typeof composeCell>) => st.visual?.colorScale?.argb ?? st.properties['fill.bg']
        withExactCfEvidence([], () => {
          const st = composeCell(rules, outs)
          expect(painted(st), `${order} ${cap}`).toBeDefined()
          expect(st.visual?.iconSet, `${order} ${cap}`).toBeUndefined()
          expect(st.unresolvedWrites?.has('iconSet'), `${order} ${cap}`).toBe(true)
        })
        withExactCfEvidence([cap], () => {
          const st = composeCell(rules, outs)
          expect(painted(st), `${order} ${cap}`).toBeDefined()
          expect(st.visual?.iconSet?.iconId, `${order} ${cap}`).toBe(2)
        })
      }
    }
    // A font-only DXF never gates the icon; a non-matching DXF fill neither.
    withExactCfEvidence([], () => {
      const font = composeCell([dxfRule(1, 1, { 'font.bold': true }), iconRule(2, 2)], [yes, iconOut])
      expect(font.properties['font.bold']).toBe(true)
      expect(font.visual?.iconSet).toBeDefined()
      expect(composeCell([dxfRule(1, 1, RED_FILL), iconRule(2, 2)], [no, iconOut]).visual?.iconSet).toBeDefined()
    })
  })

  it('a matching p1 DXF stopIfTrue → no icon', () => {
    for (const evidence of [[], ['icon:composition-matching-stopIfTrue-ahead', 'icon:composition-dxf-fill']]) {
      withExactCfEvidence(evidence, () => {
        const st = composeCell([dxfRule(1, 1, { 'font.bold': true }, true), iconRule(2, 2)], [yes, iconOut])
        expect(st.visual?.iconSet).toBeUndefined()
        expect(st.properties['font.bold']).toBe(true)
      })
    }
    withExactCfEvidence([], () => {
      expect(composeCell([dxfRule(1, 1, { 'font.bold': true }, true), iconRule(2, 2)], [no, iconOut]).visual?.iconSet).toBeDefined()
    })
  })

  it('an icon with stopIfTrue="1" → tainted contains ALL until icon:composition-icon-stopIfTrue', () => {
    const rules = [iconRule(1, 1, true), dxfRule(2, 2, { ...RED_FILL, 'font.bold': true })]
    for (const outcome of [iconOut, no, unknown]) {
      withExactCfEvidence([], () => {
        const st = composeCell(rules, [outcome, yes])
        expect(st.unresolvedWrites?.has('ALL')).toBe(true)
        expect(st.properties).toEqual({})
        expect(st.visual?.iconSet).toBeUndefined()
      })
    }
    withExactCfEvidence(['icon:composition-icon-stopIfTrue', 'icon:composition-dxf-fill'], () => {
      const st = composeCell(rules, [iconOut, yes])
      expect(st.unresolvedWrites).toBeUndefined()
      expect(st.properties['font.bold']).toBe(true)
      expect(st.visual?.iconSet?.ruleId).toBe(1)
    })
  })

  it('a NoIcons win over a lower matching icon → the lower icon is withheld until icon:composition-noicons-slot', () => {
    const rules = [iconRule(1, 1), iconRule(2, 2)]
    for (const evidence of [['icon:composition-overlapping-icons'], ['icon:composition-overlapping-icons', 'icon:composition-noicons-slot']]) {
      withExactCfEvidence(evidence, () => {
        const st = composeCell(rules, [ready(noIcons), iconOut])
        expect(st.visual?.iconSet).toBeUndefined()
        expect(st.unresolvedWrites?.has('iconSet')).toBe(true)
      })
    }
    // With the capability AND its selected reading (NoIcons occupies the slot), the NoIcons bucket wins outright.
    withExactCfEvidence(['icon:composition-overlapping-icons', 'icon:composition-noicons-slot'], () => {
      injectCfIconSelections({ 'icon:composition-noicons-slot': 'first-paints' })
      const st = composeCell(rules, [ready(noIcons), iconOut])
      expect(st.visual?.iconSet).toEqual({ ruleId: 1, ...noIcons })
      expect(st.unresolvedWrites).toBeUndefined()
      injectCfIconSelections({ 'icon:composition-noicons-slot': 'second-paints' })
      const lower = composeCell(rules, [ready(noIcons), iconOut])
      expect(lower.visual?.iconSet).toEqual({ ruleId: 2, ...icon })
    })
    // A lone NoIcons bucket is a valid descriptor (it paints nothing; Task 29 decides inset and text).
    withExactCfEvidence([], () => {
      expect(composeCell([iconRule(1, 1)], [ready(noIcons)]).visual?.iconSet).toEqual({ ruleId: 1, ...noIcons })
    })
  })

  it('an unresolved p1 stopIfTrue stub → no icon', () => {
    withExactCfEvidence(['icon:composition-unresolved-stub', 'icon:composition-dxf-fill'], () => {
      const st = composeCell([unresolvedStub(1, 1, true), iconRule(2, 2)], [unknown, iconOut])
      expect(st.visual?.iconSet).toBeUndefined()
      expect(st.unresolvedWrites?.has('ALL')).toBe(true)
    })
  })
})
