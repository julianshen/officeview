import { describe, expect, it } from 'vitest'
import { render } from '@testing-library/react'
import { GridSemantics } from '../src/components/xlsx/GridSemantics'
import { gridFixture } from './helpers/xlsx-grid-fixture'
import {
  SemanticPresentationCache,
  withCfDescription,
} from '../src/components/xlsx/formatPresentation'
import { prepareCfTextScope } from '../src/xlsx/cf/text'
import { describeCfWin } from '../src/xlsx/cf/describe'
import type { XlsxCell, XlsxSheet } from '../src/xlsx/types'
import type { FormatBudget } from '../src/xlsx/view/formatWorkspace'
import { FindError } from '../src/xlsx/view/find'
import { withInjectedCfEvidence } from '../src/xlsx/cf/evidence'

const SCALE_EVIDENCE = [
  'scale:interpolation-color-space',
  'scale:interpolation-rounding-ladder',
  'scale:thresholds-min-max',
]

function trackingBudget(refuse = false): { budget: FormatBudget; net: () => number; settled: object[] } {
  let held = 0
  const retained = new Map<object, number>()
  const settled: object[] = []
  const budget: FormatBudget = {
    reserve: bytes => { held += bytes; if (refuse) { held -= bytes; throw new FindError('resource-limit', 'refused') } },
    release: bytes => { held -= bytes },
    retain: (targets, bytes, commit) => { for (const target of targets) retained.set(target, (retained.get(target) ?? 0) + bytes); commit() },
    settle: target => { settled.push(target); held -= retained.get(target) ?? 0; retained.delete(target) },
  }
  return { budget, net: () => held, settled }
}

function scaleA11ySheet(locale = 'zh-TW'): XlsxSheet {
  const cells: XlsxCell[] = [
    { ref: 'A1', row: 0, col: 0, styleIndex: 0, value: 0, hasCachedValue: true },
    { ref: 'A2', row: 1, col: 0, styleIndex: 0, value: 72, hasCachedValue: true },
    { ref: 'A3', row: 2, col: 0, styleIndex: 0, value: 100, hasCachedValue: true },
  ]
  return {
    name: 'A11ySheet',
    sheetId: 'a11y',
    semantics: { locale, dateSystem: '1900' },
    rows: [
      { index: 0, cells: [cells[0]!] },
      { index: 1, cells: [cells[1]!] },
      { index: 2, cells: [cells[2]!] },
    ],
    cols: [],
    merges: [],
    mergeRanges: [],
    conditionalFormatting: {
      diagnostics: [],
      rules: [
        {
          id: 1,
          priority: 1,
          documentOrder: 0,
          stopIfTrue: false,
          ranges: [{ firstRow: 0, lastRow: 2, firstCol: 0, lastCol: 0 }],
          writes: new Set(['colorScale', 'fill']),
          definition: {
            kind: 'colorScale',
            points: [
              { bound: { type: 'min' }, argb: 0xfff8696b },
              { bound: { type: 'max' }, argb: 0xff63be7b },
            ],
            dialect: 'base',
            formulas: [],
          },
        },
      ],
    },
  } as unknown as XlsxSheet
}

describe('Task 16 — Colour scale accessibility and descriptors (OD-2)', () => {
  it('direct describeCfWin produces bounded zh-TW and en labels with position in range', () => {
    // zh-TW: 「色階 72%」
    const zhDesc = describeCfWin({}, undefined, undefined, 'zh-TW', { fraction: 0.72 })
    expect(zhDesc).toBeDefined()
    expect(zhDesc?.text).toBe('色階 72%')
    expect(zhDesc?.key).toBe('cf:colorScale')
    expect(zhDesc!.text.length).toBeLessThanOrEqual(128)
    expect(zhDesc!.text).not.toContain('F8696B')
    expect(zhDesc!.text).not.toContain('63BE7B')
    expect(zhDesc!.text).not.toContain('#')

    // en-US: "Colour scale, 72% of range"
    const enDesc = describeCfWin({}, undefined, undefined, 'en-US', { fraction: 0.72 })
    expect(enDesc).toBeDefined()
    expect(enDesc?.text).toBe('Colour scale, 72% of range')
    expect(enDesc?.key).toBe('cf:colorScale')
    expect(enDesc!.text.length).toBeLessThanOrEqual(128)

    // Bounded to 0-100:
    const below = describeCfWin({}, undefined, undefined, 'zh-TW', { fraction: -0.5 })
    expect(below?.text).toBe('色階 0%')
    const above = describeCfWin({}, undefined, undefined, 'en-US', { fraction: 1.5 })
    expect(above?.text).toBe('Colour scale, 100% of range')

    // An unresolved scale announces nothing scale-specific
    const unevidenced = describeCfWin({}, new Set(['colorScale']), undefined, 'zh-TW')
    expect(unevidenced?.text).not.toContain('色階')
    const unresolvedReason = describeCfWin({}, undefined, 'cf-unsupported-scale', 'zh-TW')
    expect(unresolvedReason?.text).not.toContain('色階')
  })

  it('scale cell announces a bounded zh-TW label with its position in the range, without a tab stop', async () => {
    await withInjectedCfEvidence(SCALE_EVIDENCE, async () => {
      const sheet = scaleA11ySheet('zh-TW')
      const { budget, settled } = trackingBudget()

      const scope = await prepareCfTextScope(sheet, {
        rects: [{ r0: 0, r1: 2, c0: 0, c1: 0 }],
        signal: new AbortController().signal,
        locale: 'zh-TW',
        budget,
      })
      expect(scope).toBeDefined()

      // A2 is value 72 between min 0 and max 100 -> position 72%
      const desc = scope!.describe(1, 0)
      expect(desc).toBeDefined()
      expect(desc?.text).toBe('色階 72%')
      expect(desc?.key).toBe('cf:colorScale')

      // Grid semantics integration: accessible name carries the description
      const cell = sheet.rows[1]!.cells[0]!
      const name = withCfDescription('72', desc?.text)
      expect(name).toBe('72; 色階 72%')

      const ui = render(
        <GridSemantics
          {...gridFixture({
            cellPresentation: () => ({ text: name }),
          })}
        />,
      )
      try {
        const cells = ui.getAllByRole('gridcell', { name: /色階 72%/ })
        expect(cells.length).toBeGreaterThan(0)
        expect(cells[0]!.textContent).toContain('72')

        // Keyboard focus order is unchanged: only the grid itself is tabbable (no extra tab stop)
        const tabbable = [...ui.container.querySelectorAll('[tabindex]')].filter(
          node => node.getAttribute('tabindex') !== '-1',
        )
        expect(tabbable).toHaveLength(1)
        expect(tabbable[0]!.getAttribute('role')).toBe('grid')
      } finally {
        ui.unmount()
      }

      // SemanticPresentationCache: changing the scale generation drops the stale
      // semantic entry and settles its charged strings
      const cache = new SemanticPresentationCache(sheet, budget)
      const inputG1 = { gen: 1, style: undefined, descriptorKey: desc?.key, description: desc?.text }
      const presG1 = cache.get(cell, inputG1)
      expect(presG1.description).toBe('色階 72%')

      // Generation 2 with new rule version
      const inputG2 = { gen: 2, style: undefined, descriptorKey: desc?.key, description: desc?.text }
      const presG2 = cache.get(cell, inputG2)
      expect(presG2.description).toBe('色階 72%')

      // Settled old entry when generation changed
      expect(settled.length).toBeGreaterThan(0)

      cache.dispose()
      scope!.dispose()
    })
  })
})
