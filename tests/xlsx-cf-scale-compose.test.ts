import { describe, expect, it } from 'vitest'
import type { CfRule, CfOutcome, CfScalePoint } from '../src/xlsx/cf/model'
import type { XlsxCellStyle } from '../src/xlsx/types'
import * as dxf from '../src/xlsx/cf/dxf'
import * as composition from '../src/xlsx/cf/compose'
import { withExactCfEvidence } from '../src/xlsx/cf/evidence'
import stage2CompositionFixture from './fixtures/xlsx-cf/refs/stage2-scale-composition.json'

const compose = composition.composeCell
const baseStyle = (fillColor = 'FFFFFFFF'): XlsxCellStyle => ({ numFmtId: 0, fillColor })

function getObservedArgb(ref: string): string {
  const cell = stage2CompositionFixture.observedCells.find(c => c.ref === ref)
  if (!cell || cell.fillSource !== 'excel-display-format') {
    throw new Error(`Missing trusted observed display-format cell ${ref}`)
  }
  return cell.fillRgb
}

// Trusted Excel display format ARGB values from stage2-scale-composition.json observedCells
const ARGB_B2 = getObservedArgb('B2') // FFFFC7CE (DXF red fill)
const ARGB_B4 = getObservedArgb('B4') // FFDB7A6E (scale)
const ARGB_C2 = getObservedArgb('C2') // FF63BE7B (scale)
const ARGB_D2 = getObservedArgb('D2') // FF63BE7B (scale)
const ARGB_F2 = getObservedArgb('F2') // FFFFFFFF (halted / base)
const ARGB_F4 = getObservedArgb('F4') // FFDB7A6E (scale)
const ARGB_H2 = getObservedArgb('H2') // FF63BE7B (scale)
const ARGB_J2 = getObservedArgb('J2') // FFFFC7CE (lower DXF fill on text)

function dxfRule(
  id: number,
  priority: number,
  docOrder: number,
  properties: Record<string, string | number | boolean | null>,
  writes: string[] = Object.keys(properties),
  stop = false,
): CfRule {
  return {
    id,
    priority,
    documentOrder: docOrder,
    ranges: [{ firstRow: 0, lastRow: 10, firstCol: 0, lastCol: 10 }],
    stopIfTrue: stop,
    writes: new Set(writes),
    dxf: properties,
    definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 0 }], dialect: 'base', formulas: [] },
  }
}

function scaleRule(
  id: number,
  priority: number,
  docOrder: number,
  stop = false,
  points?: readonly [CfScalePoint, CfScalePoint],
): CfRule {
  const pts: readonly [CfScalePoint, CfScalePoint] = points ?? [
    { bound: { type: 'min' }, argb: 0xfff8696b },
    { bound: { type: 'max' }, argb: 0xff63be7b },
  ]
  return {
    id,
    priority,
    documentOrder: docOrder,
    ranges: [{ firstRow: 0, lastRow: 10, firstCol: 0, lastCol: 10 }],
    stopIfTrue: stop,
    writes: new Set(['colorScale', 'fill']),
    definition: { kind: 'colorScale', points: pts, dialect: 'base', formulas: [] },
  }
}

const yes = { kind: 'ready', value: true } as const
const no = { kind: 'no-match' } as const
const unknown = { kind: 'unresolved', reason: 'cf-unsupported-rule' } as const

describe('CF Stage 2 Task 15 — Scale fill-slot composition (S2-I, S2-J)', () => {
  describe('unknown fill order withholds both fills', () => {
    it('with no composition evidence, DXF red fill and scale (either order, or equal priority) produce no CF fill, while DXF font still applies', async () => {
      const redDxf = dxfRule(1, 1, 0, { 'fill.pattern': 'solid', 'fill.bg': 'FFFFC7CE', 'font.bold': true }, ['fill', 'font.bold'])
      const scale = scaleRule(2, 2, 1)

      const scaleOutcome: CfOutcome<number> = { kind: 'ready', value: parseInt(ARGB_C2, 16) }

      // 1. DXF p1, scale p2 — with NO evidence: neither fill paints, font still applies
      await withExactCfEvidence([], async () => {
        const state = await compose([redDxf, scale], [yes, scaleOutcome])
        expect(state.properties['font.bold']).toBe(true)
        expect(state.properties['fill.bg']).toBeUndefined()
        expect(state.properties['fill.pattern']).toBeUndefined()
        expect(state.visual?.colorScale).toBeUndefined()
        expect(state.unresolvedWrites?.has('fill')).toBe(true)

        // Applied to cell style: font bold is applied, fillColor remains base
        const styled = dxf.resolveCfCellStyle(baseStyle('FFFFFFFF'), state)
        expect(styled.bold).toBe(true)
        expect(styled.fillColor).toBe('FFFFFFFF')
      })

      // 2. Scale p1, DXF p2 — with NO evidence: neither fill paints, font still applies
      const scaleP1 = scaleRule(1, 1, 0)
      const redDxfP2 = dxfRule(2, 2, 1, { 'fill.pattern': 'solid', 'fill.bg': 'FFFFC7CE', 'font.bold': true }, ['fill', 'font.bold'])
      await withExactCfEvidence([], async () => {
        const state = await compose([scaleP1, redDxfP2], [scaleOutcome, yes])
        expect(state.properties['font.bold']).toBe(true)
        expect(state.properties['fill.bg']).toBeUndefined()
        expect(state.visual?.colorScale).toBeUndefined()
        expect(state.unresolvedWrites?.has('fill')).toBe(true)

        const styled = dxf.resolveCfCellStyle(undefined, state)
        expect(styled.bold).toBe(true)
        expect(styled.fillColor).toBeUndefined()
      })

      // 3. Equal priority (priority 5, docOrder 0 and 1) — with NO evidence: neither fill paints
      const dxfEqual = dxfRule(1, 5, 0, { 'fill.pattern': 'solid', 'fill.bg': 'FFFFC7CE', 'font.bold': true }, ['fill', 'font.bold'])
      const scaleEqual = scaleRule(2, 5, 1)
      await withExactCfEvidence([], async () => {
        const state = await compose([dxfEqual, scaleEqual], [yes, scaleOutcome])
        expect(state.properties['font.bold']).toBe(true)
        expect(state.properties['fill.bg']).toBeUndefined()
        expect(state.visual?.colorScale).toBeUndefined()
        expect(state.unresolvedWrites?.has('fill')).toBe(true)
      })
    })

    it('under each measured stage2-scale-composition variant, only the measured winner paints', async () => {
      const redDxf = dxfRule(1, 1, 0, { 'fill.pattern': 'solid', 'fill.bg': ARGB_B2 }, ['fill'])
      const scale = scaleRule(2, 2, 1)
      const scaleOutcomeB2: CfOutcome<number> = { kind: 'ready', value: parseInt(getObservedArgb('C2'), 16) }
      const scaleOutcomeB4: CfOutcome<number> = { kind: 'ready', value: parseInt(ARGB_B4, 16) }

      // Case 1: scale:composition-dxf-p1-scale-p2
      await withExactCfEvidence(['scale:composition-dxf-p1-scale-p2'], async () => {
        // When DXF matches: DXF wins (observed cell B2)
        const stateMatch = await compose([redDxf, scale], [yes, scaleOutcomeB2])
        expect(stateMatch.properties['fill.bg']).toBe(ARGB_B2)
        expect(stateMatch.visual?.colorScale).toBeUndefined()
        const styleMatch = dxf.resolveCfCellStyle(undefined, stateMatch)
        expect(styleMatch.fillColor).toBe(ARGB_B2)

        // When DXF does not match: scale at p2 wins (observed cell B4)
        const stateNoMatch = await compose([redDxf, scale], [no, scaleOutcomeB4])
        expect(stateNoMatch.properties['fill.bg']).toBeUndefined()
        expect(stateNoMatch.visual?.colorScale).toEqual({ ruleId: 2, argb: parseInt(ARGB_B4, 16) })
        const styleNoMatch = dxf.resolveCfCellStyle(undefined, stateNoMatch)
        expect(styleNoMatch.fillColor).toBe(ARGB_B4)
      })

      // Case 2: scale:composition-scale-p1-dxf-p2
      const scaleP1 = scaleRule(1, 3, 0)
      const greenDxfP2 = dxfRule(2, 4, 1, { 'fill.pattern': 'solid', 'fill.bg': 'FFC6EFCE' }, ['fill'])
      const scaleOutcomeC2: CfOutcome<number> = { kind: 'ready', value: parseInt(ARGB_C2, 16) }
      await withExactCfEvidence(['scale:composition-scale-p1-dxf-p2'], async () => {
        // When both match: scale at p1 wins over DXF at p2 (observed cell C2)
        const state = await compose([scaleP1, greenDxfP2], [scaleOutcomeC2, yes])
        expect(state.properties['fill.bg']).toBeUndefined()
        expect(state.visual?.colorScale).toEqual({ ruleId: 1, argb: parseInt(ARGB_C2, 16) })
        const style = dxf.resolveCfCellStyle(undefined, state)
        expect(style.fillColor).toBe(ARGB_C2)
      })

      // Case 3: scale:composition-equal-priority
      const dxfEqual = dxfRule(1, 5, 0, { 'fill.pattern': 'solid', 'fill.bg': 'FFFFC7CE' }, ['fill'])
      const scaleEqual = scaleRule(2, 5, 1)
      const scaleOutcomeD2: CfOutcome<number> = { kind: 'ready', value: parseInt(ARGB_D2, 16) }
      await withExactCfEvidence(['scale:composition-equal-priority'], async () => {
        // When both match: in Excel display format (observed cell D2), scale won!
        const state = await compose([dxfEqual, scaleEqual], [yes, scaleOutcomeD2])
        expect(state.properties['fill.bg']).toBeUndefined()
        expect(state.visual?.colorScale).toEqual({ ruleId: 2, argb: parseInt(ARGB_D2, 16) })
        const style = dxf.resolveCfCellStyle(undefined, state)
        expect(style.fillColor).toBe(ARGB_D2)
      })
    })

    it('two overlapping scales withhold the fill', async () => {
      const scale1 = scaleRule(1, 14, 0)
      const scale2 = scaleRule(2, 15, 1, false, [
        { bound: { type: 'min' }, argb: 0xff0000ff },
        { bound: { type: 'max' }, argb: 0xffffff00 },
      ])
      const outcome1: CfOutcome<number> = { kind: 'ready', value: 0xff63be7b }
      const outcome2: CfOutcome<number> = { kind: 'ready', value: 0xffffff00 }

      // With or without evidence: two resolved scales on one cell withhold fill (observed cell I2 FFFFFFFF)
      await withExactCfEvidence(['scale:composition-overlapping-scales'], async () => {
        const state = await compose([scale1, scale2], [outcome1, outcome2])
        expect(state.visual?.colorScale).toBeUndefined()
        expect(state.properties['fill.bg']).toBeUndefined()
        expect(state.unresolvedWrites?.has('fill')).toBe(true)

        const style = dxf.resolveCfCellStyle(baseStyle('FFFFFFFF'), state)
        expect(style.fillColor).toBe('FFFFFFFF')
      })
    })

    it('font-only DXF merges cleanly with scale fill', async () => {
      const fontDxf = dxfRule(1, 6, 0, { 'font.bold': true, 'font.color': 'FF006100' }, ['font.bold', 'font.color'])
      const scale = scaleRule(2, 7, 1)
      const scaleOutcomeE2: CfOutcome<number> = { kind: 'ready', value: parseInt(ARGB_C2, 16) }

      await withExactCfEvidence(['scale:composition-font-only-dxf'], async () => {
        const state = await compose([fontDxf, scale], [yes, scaleOutcomeE2])
        expect(state.properties['font.bold']).toBe(true)
        expect(state.properties['font.color']).toBe('FF006100')
        expect(state.visual?.colorScale).toEqual({ ruleId: 2, argb: parseInt(ARGB_C2, 16) })

        const style = dxf.resolveCfCellStyle(undefined, state)
        expect(style.bold).toBe(true)
        expect(style.color).toBe('FF006100')
        expect(style.fillColor).toBe(ARGB_C2)
      })
    })
  })

  describe('stop rules and scales', () => {
    it('an unresolved p1 stopIfTrue stub withholds a resolved p2 scale', async () => {
      const stopStub = dxfRule(1, 1, 0, {}, ['fill'], true) // stopIfTrue: true
      const scale = scaleRule(2, 2, 1)
      const scaleOutcome: CfOutcome<number> = { kind: 'ready', value: parseInt(ARGB_C2, 16) }

      // When p1 stub is unresolved and carries stopIfTrue: withholds p2 scale
      const state = await compose([stopStub, scale], [unknown, scaleOutcome])
      expect(state.visual?.colorScale).toBeUndefined()
      expect(state.unresolvedWrites?.has('ALL')).toBe(true)
    })

    it('a matching p1 DXF stopIfTrue halts the scale only under evidence', async () => {
      const stopDxf = dxfRule(1, 8, 0, {}, ['fill'], true) // bare stopIfTrue rule
      const scale = scaleRule(2, 9, 1)
      const scaleOutcomeF2: CfOutcome<number> = { kind: 'ready', value: parseInt(ARGB_C2, 16) }
      const scaleOutcomeF4: CfOutcome<number> = { kind: 'ready', value: parseInt(ARGB_F4, 16) }

      // Without evidence: matching stopIfTrue halts unverified -> fill withheld/unresolved
      await withExactCfEvidence([], async () => {
        const state = await compose([stopDxf, scale], [yes, scaleOutcomeF2])
        expect(state.visual?.colorScale).toBeUndefined()
        expect(state.unresolvedWrites?.has('fill')).toBe(true)
      })

      // Under evidence scale:composition-matching-stopIfTrue-ahead:
      await withExactCfEvidence(['scale:composition-matching-stopIfTrue-ahead'], async () => {
        // Matching p1 stopIfTrue halts the scale: scale does NOT paint (observed cell F2 FFFFFFFF)
        const stateHalted = await compose([stopDxf, scale], [yes, scaleOutcomeF2])
        expect(stateHalted.visual?.colorScale).toBeUndefined()
        expect(stateHalted.unresolvedWrites?.has('fill') ?? false).toBe(false)
        const styleHalted = dxf.resolveCfCellStyle(baseStyle('FFFFFFFF'), stateHalted)
        expect(styleHalted.fillColor).toBe(ARGB_F2)

        // Non-matching p1 stopIfTrue does NOT halt the scale (observed cell F4)
        const statePassed = await compose([stopDxf, scale], [no, scaleOutcomeF4])
        expect(statePassed.visual?.colorScale).toEqual({ ruleId: 2, argb: parseInt(ARGB_F4, 16) })
        const stylePassed = dxf.resolveCfCellStyle(undefined, statePassed)
        expect(stylePassed.fillColor).toBe(ARGB_F4)
      })
    })

    it('a scale carrying stopIfTrue="1" is unresolved until composition variant lands (S2-J); per OD-1, once measured, the flag is ignored', async () => {
      const scaleWithStop = scaleRule(1, 12, 0, true) // stopIfTrue = true
      const lowerDxf = dxfRule(2, 13, 1, { 'fill.pattern': 'solid', 'fill.bg': 'FFFFC7CE', 'font.bold': true }, ['fill', 'font.bold'])
      const scaleOutcome: CfOutcome<number> = { kind: 'ready', value: parseInt(ARGB_H2, 16) }

      // 1. BEFORE evidence: scale with stopIfTrue="1" is unresolved, withholds p2 DXF fill AND font (S2-J)
      await withExactCfEvidence([], async () => {
        const state = await compose([scaleWithStop, lowerDxf], [scaleOutcome, yes])
        expect(state.visual?.colorScale).toBeUndefined()
        expect(state.properties['font.bold']).toBeUndefined()
        expect(state.properties['fill.bg']).toBeUndefined()
        expect(state.unresolvedWrites?.has('ALL')).toBe(true)
      })

      // 2. AFTER evidence scale:composition-scale-stopIfTrue (per OD-1):
      // stopIfTrue flag is ignored on scale; scale resolves; S2-I governs fill slot; lower font applies
      await withExactCfEvidence(['scale:composition-scale-stopIfTrue', 'scale:composition-scale-p1-dxf-p2'], async () => {
        const state = await compose([scaleWithStop, lowerDxf], [scaleOutcome, yes])
        // Scale resolves and wins fill slot
        expect(state.visual?.colorScale).toEqual({ ruleId: 1, argb: parseInt(ARGB_H2, 16) })
        // Lower rule is NOT halted by stopIfTrue: font.bold applies!
        expect(state.properties['font.bold']).toBe(true)
        expect(state.properties['fill.bg']).toBeUndefined() // fill slot was won by higher scale

        const style = dxf.resolveCfCellStyle(undefined, state)
        expect(style.bold).toBe(true)
        expect(style.fillColor).toBe(ARGB_H2)
      })
    })

    it('a text cell inside the scale range taints the fill for a lower DXF until measured (S2-G)', async () => {
      const scale = scaleRule(1, 16, 0)
      const lowerDxf = dxfRule(2, 17, 1, { 'fill.pattern': 'solid', 'fill.bg': ARGB_J2, 'font.bold': true }, ['fill', 'font.bold'])
      // For a text cell: scale evaluation produces { kind: 'no-match' } or text-non-participant outcome
      const textScaleOutcome: CfOutcome<number> = { kind: 'no-match' }

      // Without evidence scale:composition-text-cell-lower-dxf: text cell taints fill for lower DXF
      await withExactCfEvidence([], async () => {
        const state = await compose([scale, lowerDxf], [textScaleOutcome, yes], {
          isTextCellInScale: true,
        } as any)
        expect(state.properties['font.bold']).toBe(true)
        expect(state.properties['fill.bg']).toBeUndefined()
        expect(state.unresolvedWrites?.has('fill')).toBe(true)
      })

      // With evidence scale:composition-text-cell-lower-dxf: lower DXF fill paints (observed cell J2 FFFFC7CE)
      await withExactCfEvidence(['scale:composition-text-cell-lower-dxf'], async () => {
        const state = await compose([scale, lowerDxf], [textScaleOutcome, yes], {
          isTextCellInScale: true,
        } as any)
        expect(state.properties['font.bold']).toBe(true)
        expect(state.properties['fill.bg']).toBe(ARGB_J2)
        const style = dxf.resolveCfCellStyle(undefined, state)
        expect(style.bold).toBe(true)
        expect(style.fillColor).toBe(ARGB_J2)
      })
    })
  })
})
