/**
 * Ruling S3-A11Y-PAGED (docs/validation/xlsx-cf.md §11): the forced-paged view's grid semantics read a wire scope that
 * never paints. Given the frame's own host source (each cell's logical box from the sheet metrics and its screen base
 * style), that scope decides whether a bar paints with the painter's own predicate (`barWithheld` at the capture
 * scale), so it announces exactly what the resident and worker legs announce: the bar descriptor where the bar paints,
 * "unavailable" where painting withholds it, and never a bar the frame would not paint.
 */
import { describe, expect, it } from 'vitest'
import { prepareCfWireScope, type CfFrameHostSource } from '../src/xlsx/cf/text'
import { createCfMemory, createCfWorkBudget } from '../src/xlsx/cf/limits'
import { residentSource } from '../src/xlsx/cf/render'
import { cfRulesVersion, prepareCfRemote } from '../src/xlsx/view/pagedSource'
import type { CfStamp } from '../src/xlsx/cf/model'
import type { CfReadResult } from '../src/xlsx/view/contracts'
import type { XlsxCellStyle, XlsxSheet } from '../src/xlsx/types'
import { barSheet, withBars } from './helpers/cf-bars'

const UNAVAILABLE = { key: 'cf-unavailable', text: 'Conditional formatting unavailable' }

const wireOf = async (sheet: XlsxSheet, rows: number): Promise<CfReadResult> => {
  const rules = sheet.conditionalFormatting!, sheetId = sheet.sheetId!
  const st: CfStamp = { owner: 'paged', sheetId, workbookValueGen: 1, cfGen: cfRulesVersion(rules), viewGen: 1, structureGen: 1, clockGen: 1 }
  const memory = createCfMemory({})
  const wire = await prepareCfRemote({ rects: [{ r0: 0, r1: rows - 1, c0: 0, c1: 0 }], stamp: { ...st }, grantBytes: 1 << 22 }, sheetId, residentSource(sheet, st, memory, 'engine'), rules, memory, { budget: () => createCfWorkBudget({ now: () => 0 }) })
  expect(wire.kind).toBe('ready')
  return wire as CfReadResult
}
const scopeOver = (wire: CfReadResult, rows: number, locale: string, hosts?: CfFrameHostSource) =>
  prepareCfWireScope(async () => ({ result: structuredClone(wire), release() {} }), [{ r0: 0, r1: rows - 1, c0: 0, c1: 0 }],
    { budget: { reserve() {}, release() {} }, locale, ...(hosts ? { hosts } : {}) })
/** The frame's host source: a `width` × 20 px box per row, the given screen base style. */
const frameHosts = (width = 200, base?: XlsxCellStyle): CfFrameHostSource => ({
  cell: (row, col) => ({ box: { x: col * width, y: row * 20, width, height: 20 }, base }),
})

describe('the forced-paged semantic scope decides paintability from the frame host (S3-A11Y-PAGED)', () => {
  it('announces a painted bar with its value position, without any paint having run (en and zh-TW)', async () => {
    await withBars(async () => {
      const wire = await wireOf(barSheet([0, 72, 100]), 3)
      const en = (await scopeOver(wire, 3, 'en-US', frameHosts()))!
      expect([0, 1, 2].map(r => en.describe(r, 0))).toEqual([
        { key: 'cf:dataBar', text: 'Data bar, 0% of range' },
        { key: 'cf:dataBar', text: 'Data bar, 72% of range' },
        { key: 'cf:dataBar', text: 'Data bar, 100% of range' },
      ])
      expect(en.barAt?.(1, 0)?.valueFraction).toBeCloseTo(0.72, 10)
      en.dispose()
      const zh = (await scopeOver(wire, 3, 'zh-TW', frameHosts()))!
      expect(zh.describe(1, 0)).toEqual({ key: 'cf:dataBar', text: '資料橫條 72%' })
      zh.dispose()
    })
  })

  it('a host too small to lay the bar out at the capture scale withholds it: unavailable', async () => {
    await withBars(async () => {
      const wire = await wireOf(barSheet([0, 72, 100]), 3)
      // 3 px: the 2 pt + 1 pt insets leave no span at the capture scale (commit b85ddb88).
      const scope = (await scopeOver(wire, 3, 'en-US', frameHosts(3)))!
      for (const row of [0, 1, 2]) {
        expect(scope.describe(row, 0), `row ${row}`).toEqual(UNAVAILABLE)
        expect(scope.barAt?.(row, 0), `row ${row}`).toBeUndefined()
      }
      scope.dispose()
    })
  })

  it('a cell the frame has no exact host for (unknown, hidden, outside the metrics) withholds: unavailable', async () => {
    await withBars(async () => {
      const wire = await wireOf(barSheet([0, 72, 100]), 3)
      const scope = (await scopeOver(wire, 3, 'en-US', { cell: (row, col) => row === 1 ? undefined : frameHosts().cell(row, col) }))!
      expect(scope.describe(1, 0)).toEqual(UNAVAILABLE)
      expect(scope.describe(2, 0)).toEqual({ key: 'cf:dataBar', text: 'Data bar, 100% of range' })
      scope.dispose()
    })
  })

  it('a non-white axis over the host fill the frame paints withholds: unavailable', async () => {
    await withBars(async () => {
      const wire = await wireOf(barSheet([-50, 50, 100]), 3)
      const filled = (await scopeOver(wire, 3, 'en-US', frameHosts(200, { numFmtId: 0, fillColor: 'FFFFFF00' })))!
      expect(filled.describe(1, 0)).toEqual(UNAVAILABLE)
      filled.dispose()
      const plain = (await scopeOver(wire, 3, 'en-US', frameHosts()))!
      expect(plain.describe(1, 0)).toEqual({ key: 'cf:dataBar', text: 'Data bar, 67% of range' })
      plain.dispose()
    })
  })

  it('unverified paint capabilities withhold the bar on the paged path: unavailable', async () => {
    for (const options of [{ paint: false }, { omit: ['bar:geometry-narrow-clip'] }] as const) {
      await withBars(async () => {
        const wire = await wireOf(barSheet([0, 72, 100]), 3)
        const scope = (await scopeOver(wire, 3, 'en-US', frameHosts()))!
        expect(scope.describe(1, 0), JSON.stringify(options)).toEqual(UNAVAILABLE)
        expect(scope.barAt?.(1, 0)).toBeUndefined()
        scope.dispose()
      }, options)
    }
  })

  it('agrees with what the painter does: each hosted descriptor equals a host-less scope\'s after painting the same hosts', async () => {
    const { createCanvas } = await import('canvas')
    await withBars(async () => {
      const cases: ReadonlyArray<readonly [string, XlsxSheet, number, XlsxCellStyle | undefined]> = [
        ['plain', barSheet([0, 72, 100]), 200, undefined],
        ['too small', barSheet([0, 72, 100]), 3, undefined],
        ['sub-pixel share', barSheet([0, 0.001, 1000]), 200, undefined],
        ['filled axis', barSheet([-50, 50, 100]), 200, { numFmtId: 0, fillColor: 'FFFFFF00' }],
        ['axis control', barSheet([-50, 50, 100]), 200, undefined],
      ]
      for (const [name, sheet, width, base] of cases) {
        const wire = await wireOf(sheet, 3)
        const hosts = frameHosts(width, base)
        const hosted = (await scopeOver(wire, 3, 'en-US', hosts))!
        const painted = (await scopeOver(wire, 3, 'en-US'))!
        const ctx = createCanvas(width, 60).getContext('2d') as unknown as CanvasRenderingContext2D
        painted.paintBars?.(ctx, (row, col) => {
          const at = hosts.cell(row, col)!
          return { ...at.box, ...(at.base?.fillColor ? { fill: '#' + at.base.fillColor.slice(2) } : {}) }
        }, { r0: 0, c0: 0, r1: 2, c1: 0 }, { zoom: 1, dpr: 1 })
        for (const row of [0, 1, 2]) {
          expect(hosted.describe(row, 0), `${name} row ${row}`).toEqual(painted.describe(row, 0))
          expect(hosted.barAt?.(row, 0), `${name} row ${row}`).toEqual(painted.barAt?.(row, 0))
        }
        hosted.dispose(); painted.dispose()
      }
    })
  })
})
