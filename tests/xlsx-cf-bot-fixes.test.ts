// @vitest-environment node
/**
 * PR #67 bot-loop fixes (380dc376, 051c82fa, 3fc8ba38, 7acd514e): one regression test per fix, each of which fails
 * when its fix is reverted.
 */
import { describe, expect, it } from 'vitest'
import { composeCell } from '../src/xlsx/cf/compose'
import { createCfMemory, createCfWorkBudget, CF_LIMITS } from '../src/xlsx/cf/limits'
import { createCfService } from '../src/xlsx/cf/service'
import { createCfStatistics, createResidentCfSource } from '../src/xlsx/cf/statistics'
import { residentSource } from '../src/xlsx/cf/render'
import { measureCfCellBytes } from '../src/xlsx/view/contracts'
import type { CfRule, CfScalePoint, CfStamp, SourceRect } from '../src/xlsx/cf/model'
import type { XlsxCell, XlsxSheet } from '../src/xlsx/types'

const stamp: CfStamp = { owner: 'bot', sheetId: 's', workbookValueGen: 1, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1 }
const rect = (firstRow: number, lastRow: number): SourceRect => ({ firstRow, lastRow, firstCol: 0, lastCol: 0 })
const numCells = (n: number): XlsxCell[] => Array.from({ length: n }, (_, r) => ({ ref: `A${r + 1}`, row: r, col: 0, value: r + 1, styleIndex: 0 }) as XlsxCell)

describe('380dc376: a matching stopIfTrue DXF halts fill arbitration for scales below it', () => {
  const dxf = (): CfRule => ({
    id: 1, priority: 1, documentOrder: 0, stopIfTrue: true, ranges: [rect(0, 3)], writes: new Set(['fill']), dxf: { 'fill.pattern': 'solid', 'fill.bg': 'FFFFC7CE' },
    definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 0 }], dialect: 'base', formulas: [] },
  })
  const scale = (): CfRule => ({
    id: 2, priority: 2, documentOrder: 1, stopIfTrue: false, ranges: [rect(0, 3)], writes: new Set(['colorScale', 'fill']),
    definition: { kind: 'colorScale', points: [{ bound: { type: 'min' }, argb: 0xfff8696b }, { bound: { type: 'max' }, argb: 0xff63be7b }] as [CfScalePoint, CfScalePoint], dialect: 'base', formulas: [] },
  })
  it('the unreachable scale does not withhold the stopping DXF fill', () => {
    // Only the stopIfTrue-ahead evidence: DXF-p1/scale-p2 ordering was NOT measured, so a scale that took part in
    // arbitration would withhold BOTH fills. The scale is unreachable (the DXF stops), so the DXF fill must paint.
    const state = composeCell([dxf(), scale()], [{ kind: 'ready', value: true }, { kind: 'ready', value: { argb: 0xff63be7b } }],
      { hasEvidence: cap => cap === 'scale:composition-matching-stopIfTrue-ahead' })
    expect(state.properties['fill.bg']).toBe('FFFFC7CE')
    expect(state.visual?.colorScale).toBeUndefined()
  })
})

describe('051c82fa: an optional scale fraction is charged in the retained-cell accounting', () => {
  it('a scale with a fraction costs more than one without', () => {
    const base = { properties: {}, scale: { ruleId: 1, argb: 0xff63be7b } }
    expect(measureCfCellBytes({ ...base, scale: { ...base.scale, fraction: 0.5 } } as never)).toBeGreaterThan(measureCfCellBytes(base as never))
  })
})

describe('3fc8ba38 / 380dc376: sorted statistics reuse and refusals', () => {
  it('an unsorted request that reuses sorted statistics does not alias a second cache entry', async () => {
    const memory = createCfMemory({})
    const source = createResidentCfSource(numCells(50), 's', stamp, memory)
    const session = createCfStatistics(memory)
    const sorted = await session.prepareStatistics([[rect(0, 49)]], stamp, source, createCfWorkBudget({ noTimeLimit: true }), { sorted: true })
    const streaming = await session.prepareStatistics([[rect(0, 49)]], stamp, source, createCfWorkBudget({ noTimeLimit: true }), { sorted: false })
    expect(sorted[0]!.kind).toBe('ready')
    expect(streaming[0]!.kind).toBe('ready')
    expect(session.size, 'one entry under its owning (sorted) key').toBe(1)
    for (const o of [...sorted, ...streaming]) if (o.kind === 'ready') o.value.release()
    session.dispose()
    expect(memory.currentBytes).toBe(0)
  })

  it('a sorted-statistics resource refusal is not served to a later streaming request', async () => {
    const memory = createCfMemory({})
    const source = createResidentCfSource(numCells(1000), 's', stamp, memory)
    const session = createCfStatistics(memory)
    let cap = 1200
    // Find a work cap that refuses the sorted scan but admits the streaming one.
    let sortedOut
    for (; cap < 20000; cap += 200) {
      sortedOut = await session.prepareStatistics([[rect(0, 999)]], stamp, source, createCfWorkBudget({ noTimeLimit: true, maxUnits: cap }), { sorted: true })
      if (sortedOut[0]!.kind === 'unresolved') break
      for (const o of sortedOut) if (o.kind === 'ready') o.value.release()
      session.dispose()
    }
    expect(sortedOut![0]).toMatchObject({ kind: 'unresolved', reason: 'cf-resource-limit' })
    const streaming = await session.prepareStatistics([[rect(0, 999)]], stamp, source, createCfWorkBudget({ noTimeLimit: true }), { sorted: false })
    expect(streaming[0]!.kind).toBe('ready')
    for (const o of streaming) if (o.kind === 'ready') o.value.release()
    session.dispose()
  })
})

describe('7acd514e: the paint budget resumes before the scale-scope checks', () => {
  it('a long statistics scan does not exhaust the preparation budget before the merged-scope check', async () => {
    let t = 0, bumped = false
    const cells = numCells(10).map(c => Object.defineProperty({ ...c } as Record<string, unknown>, 'value', {
      get() { if (!bumped) { bumped = true; t += CF_LIMITS.activeMs + 1 } return c.value }, enumerable: true,
    })) as unknown as XlsxCell[]
    const sheet = { name: 'S', sheetId: 's', cols: [], merges: [], mergeRanges: [], rows: cells.map((cell, r) => ({ index: r, cells: [cell] })) } as unknown as XlsxSheet
    const memory = createCfMemory()
    const rules = { diagnostics: [], rules: [{
      id: 1, priority: 1, documentOrder: 0, stopIfTrue: false, ranges: [rect(0, 9)], writes: new Set(['colorScale', 'fill']),
      definition: { kind: 'colorScale', points: [{ bound: { type: 'min' }, argb: 0xfff8696b }, { bound: { type: 'max' }, argb: 0xff63be7b }], dialect: 'base', formulas: [] },
    }] } as never
    const service = createCfService(rules, memory, {
      evidence: { allows: () => true },
      budget: () => createCfWorkBudget({ now: () => t }),
      merges: [{ firstRow: 20, lastRow: 21, firstCol: 5, lastCol: 6 }, { firstRow: 30, lastRow: 31, firstCol: 5, lastCol: 6 }],
    })
    const source = residentSource(sheet, stamp, memory, 'engine')
    const lease = await service.prepareCf({ stamp, sourceRects: [rect(0, 9)], signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, source, memory)
    expect(lease.value.at(0, 0).kind).toBe('ready')
    lease.release(); service.dispose()
  })
})
