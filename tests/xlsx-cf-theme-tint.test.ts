import { describe, expect, it } from 'vitest'
import { existsSync, readFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { applyTint, resolveSheetColor, type SheetColorContext } from '../src/xlsx/sheetColors'
import { scaleColorAt } from '../src/xlsx/cf/colorScale'
import type { CfScalePoint } from '../src/xlsx/cf/model'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..')

function toHex(argb: number): string {
  return (argb >>> 0).toString(16).padStart(8, '0').toUpperCase()
}

describe('CF theme tint resolution and scale interpolation (T16-A Step 1)', () => {
  it('resolves theme 4 +0.4 and theme 5 -0.4 exactly matching Excel observedCells in stage2-scale-colors.json', () => {
    const sidecarPath = join(REPO_ROOT, 'tests/fixtures/xlsx-cf/refs/stage2-scale-colors.json')
    expect(existsSync(sidecarPath), 'stage2-scale-colors.json must exist').toBe(true)
    const sidecar = JSON.parse(readFileSync(sidecarPath, 'utf8')) as {
      observedCells: Array<{ ref: string; fillSource: string; fillRgb: string }>
    }

    const observed = new Map<string, string>()
    for (const c of sidecar.observedCells) {
      if (c.fillSource === 'excel-display-format') {
        observed.set(c.ref, c.fillRgb)
      }
    }

    // Expected values from sidecar observedCells:
    const expectedB2 = observed.get('B2') // 'FF44B3E1'
    const expectedB3 = observed.get('B3') // 'FF5997AD'
    const expectedB4 = observed.get('B4') // 'FF6E7A79'
    const expectedB5 = observed.get('B5') // 'FF835D45'
    const expectedB6 = observed.get('B6') // 'FF984010'

    expect(expectedB2).toBe('FF44B3E1')
    expect(expectedB3).toBe('FF5997AD')
    expect(expectedB4).toBe('FF6E7A79')
    expect(expectedB5).toBe('FF835D45')
    expect(expectedB6).toBe('FF984010')

    // Theme colors from stage2-scale-colors.xlsx:
    // accent1 (theme 4) = '156082'
    // accent2 (theme 5) = 'E97132'
    const context: SheetColorContext = {
      theme: {
        accent1: '156082',
        accent2: 'E97132',
      },
    }

    // Endpoints from rule:
    // Min endpoint: theme 4, tint 0.39997558519241921
    // Max endpoint: theme 5, tint -0.39997558519241921
    const tintMin = 0.39997558519241921
    const tintMax = -0.39997558519241921

    const resolvedMin = resolveSheetColor({ theme: '4', tint: String(tintMin) }, context)
    const resolvedMax = resolveSheetColor({ theme: '5', tint: String(tintMax) }, context)

    // Check endpoints resolve exactly to Excel observed B2 and B6:
    expect(resolvedMin, 'Theme 4 tint +0.4 endpoint must match Excel B2 ARGB').toBe(expectedB2)
    expect(resolvedMax, 'Theme 5 tint -0.4 endpoint must match Excel B6 ARGB').toBe(expectedB6)

    // Now check that pure colorScale evaluation interpolates B3..B5 to match Excel observed:
    const points: CfScalePoint[] = [
      { bound: { type: 'min' }, argb: parseInt(resolvedMin!, 16) },
      { bound: { type: 'max' }, argb: parseInt(resolvedMax!, 16) },
    ]
    const thresholds = [10, 90]
    const opts = {
      hasEvidence: () => true,
      colorSpace: 'sRGB' as const,
      rounding: 'truncate' as const,
    }

    const resB2 = scaleColorAt(10, points, thresholds, opts)
    const resB3 = scaleColorAt(30, points, thresholds, opts)
    const resB4 = scaleColorAt(50, points, thresholds, opts)
    const resB5 = scaleColorAt(70, points, thresholds, opts)
    const resB6 = scaleColorAt(90, points, thresholds, opts)

    expect(resB2.kind === 'ready' && toHex(resB2.value)).toBe(expectedB2)
    expect(resB3.kind === 'ready' && toHex(resB3.value)).toBe(expectedB3)
    expect(resB4.kind === 'ready' && toHex(resB4.value)).toBe(expectedB4)
    expect(resB5.kind === 'ready' && toHex(resB5.value)).toBe(expectedB5)
    expect(resB6.kind === 'ready' && toHex(resB6.value)).toBe(expectedB6)
  })
})

describe('applyTint reproduces the canonical Excel palette from its STORED tint strings (M3)', () => {
  // Excel stores 0.4 / 0.6 / 0.8 / -0.25 / -0.5 with quantisation error, so the +0.01 luminance epsilon in applyTint
  // is load-bearing: a "cleanup" to a 1e-9 epsilon moves sixteen of these swatches by 1-2 per channel.
  const STORED = ['0.79998168889431442', '0.59999389629810485', '0.39997558519241921', '-0.249977111117893', '-0.499984740745262'].map(Number)
  const SWATCHES: Record<string, string[]> = {
    ED7D31: ['FCE4D6', 'F8CBAD', 'F4B084', 'C65911', '833C0C'],
    FFC000: ['FFF2CC', 'FFE699', 'FFD966', 'BF8F00', '806000'],
    A5A5A5: ['EDEDED', 'DBDBDB', 'C9C9C9', '7B7B7B', '525252'],
    '70AD47': ['E2EFDA', 'C6E0B4', 'A9D08E', '548235', '375623'],
    '4472C4': ['D9E1F2', 'B4C6E7', '8EA9DB', '305496', '203764'],
    '5B9BD5': ['DDEBF7', 'BDD7EE', '9BC2E6', '2F75B5', '1F4E78'],
  }
  for (const [base, expected] of Object.entries(SWATCHES)) {
    it(`${base} with Excel's stored tints`, () => {
      expect(STORED.map(tint => applyTint(base, tint))).toEqual(expected)
    })
  }
})

