import { describe, expect, it } from 'vitest'
import { gzipSync } from 'node:zlib'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { execFileSync } from 'node:child_process'
import { builtinTableStyle, builtinSlicerStyle } from '../src/xlsx/tableStyles/builtins'
import type { ColorSpec, DxfStyle } from '../src/xlsx/tableStyles/model'

const TABLE_NAMES: string[] = [
  ...Array.from({ length: 21 }, (_, i) => `TableStyleLight${i + 1}`),
  ...Array.from({ length: 28 }, (_, i) => `TableStyleMedium${i + 1}`),
  ...Array.from({ length: 11 }, (_, i) => `TableStyleDark${i + 1}`),
]

const PIVOT_NAMES: string[] = [
  ...Array.from({ length: 28 }, (_, i) => `PivotStyleLight${i + 1}`),
  ...Array.from({ length: 28 }, (_, i) => `PivotStyleMedium${i + 1}`),
  ...Array.from({ length: 28 }, (_, i) => `PivotStyleDark${i + 1}`),
]

const SLICER_NAMES: string[] = [
  ...Array.from({ length: 6 }, (_, i) => `SlicerStyleLight${i + 1}`),
  ...Array.from({ length: 6 }, (_, i) => `SlicerStyleDark${i + 1}`),
  'SlicerStyleOther1',
  'SlicerStyleOther2',
]

const TIMELINE_NAMES: string[] = [
  ...Array.from({ length: 6 }, (_, i) => `TimeSlicerStyleLight${i + 1}`),
  ...Array.from({ length: 6 }, (_, i) => `TimeSlicerStyleDark${i + 1}`),
]

function assertValidColor(color: ColorSpec | undefined, context: string): void {
  expect(color, `${context} has color`).toBeDefined()
  if (!color) return
  const isTheme = 'theme' in color && typeof color.theme === 'number' && Number.isFinite(color.theme)
  const isRgb = 'rgb' in color && typeof color.rgb === 'string' && color.rgb.length > 0
  expect(isTheme || isRgb, `${context} color is theme or rgb`).toBe(true)
}

function assertValidDxfColors(dxf: DxfStyle, context: string): void {
  if (dxf.font?.color) assertValidColor(dxf.font.color, `${context} font`)
  if (dxf.fill?.fg) assertValidColor(dxf.fill.fg, `${context} fill.fg`)
  if (dxf.fill?.bg) assertValidColor(dxf.fill.bg, `${context} fill.bg`)
  if (dxf.border) {
    for (const [side, border] of Object.entries(dxf.border)) {
      if (border?.color) assertValidColor(border.color, `${context} border.${side}`)
    }
  }
}

describe('built-in table, pivot, slicer and timeline styles', () => {
  it('resolves all 144 table and pivot names', () => {
    expect(TABLE_NAMES).toHaveLength(60)
    expect(PIVOT_NAMES).toHaveLength(84)
    const allNames = [...TABLE_NAMES, ...PIVOT_NAMES]
    expect(allNames).toHaveLength(144)

    for (const name of TABLE_NAMES) {
      const def = builtinTableStyle(name)
      expect(def, `table style ${name} resolves`).toBeDefined()
      expect(def?.name).toBe(name)
      expect(def?.table).toBe(true)
      expect(def?.pivot).toBe(false)
      expect(Object.keys(def?.elements ?? {}).length).toBeGreaterThan(0)
    }

    for (const name of PIVOT_NAMES) {
      const def = builtinTableStyle(name)
      expect(def, `pivot style ${name} resolves`).toBeDefined()
      expect(def?.name).toBe(name)
      expect(def?.table).toBe(false)
      expect(def?.pivot).toBe(true)
      expect(Object.keys(def?.elements ?? {}).length).toBeGreaterThan(0)
    }
  })

  it('spot-checks TableStyleMedium2.headerRow and PivotStyleLight16.firstRowStripe against presetTableStyles.xml', () => {
    const medium2 = builtinTableStyle('TableStyleMedium2')
    expect(medium2).toBeDefined()
    const headerRow = medium2?.elements.headerRow?.dxf
    expect(headerRow).toBeDefined()
    expect(headerRow?.font?.bold).toBe(true)
    expect(headerRow?.font?.color).toEqual({ theme: 0 })
    expect(headerRow?.fill?.pattern).toBe('solid')
    expect(headerRow?.fill?.fg).toEqual({ theme: 4 })
    expect(headerRow?.fill?.bg).toEqual({ theme: 4 })

    const pivotLight16 = builtinTableStyle('PivotStyleLight16')
    expect(pivotLight16).toBeDefined()
    const firstRowStripe = pivotLight16?.elements.firstRowStripe?.dxf
    expect(firstRowStripe).toBeDefined()
    expect(firstRowStripe?.fill?.pattern).toBe('solid')
    const fg = firstRowStripe?.fill?.fg
    const bg = firstRowStripe?.fill?.bg
    expect(fg && 'theme' in fg ? fg.theme : undefined).toBe(0)
    expect(bg && 'theme' in bg ? bg.theme : undefined).toBe(0)
    expect(fg && 'theme' in fg ? fg.tint : undefined).toBeCloseTo(-0.15, 2)
    expect(bg && 'theme' in bg ? bg.tint : undefined).toBeCloseTo(-0.15, 2)
  })

  it('ensures every colour is theme-relative or rgb, never empty', () => {
    for (const name of [...TABLE_NAMES, ...PIVOT_NAMES]) {
      const def = builtinTableStyle(name)!
      expect(def).toBeDefined()
      for (const [elType, el] of Object.entries(def.elements)) {
        assertValidDxfColors(el.dxf, `${name}.${elType}`)
      }
    }
  })

  it('resolves slicer and timeline styles as two lookups per Ruling Q', () => {
    for (const name of SLICER_NAMES) {
      const frame = builtinTableStyle(name)
      expect(frame, `${name} tableStyle frame`).toBeDefined()
      expect(frame?.name).toBe(name)
      expect(frame?.elements.wholeTable).toBeDefined()
      expect(frame?.elements.headerRow).toBeDefined()
      assertValidDxfColors(frame!.elements.wholeTable!.dxf, `${name}.wholeTable`)
      assertValidDxfColors(frame!.elements.headerRow!.dxf, `${name}.headerRow`)

      const items = builtinSlicerStyle(name)
      expect(items, `${name} slicerStyle items`).toBeDefined()
      expect(items?.name).toBe(name)
      expect(items?.elements.selectedItemWithData).toBeDefined()
      expect(items?.elements.unselectedItemWithData).toBeDefined()
      expect(items?.elements.selectedItemWithNoData).toBeDefined()
      expect(items?.elements.unselectedItemWithNoData).toBeDefined()
      expect(items?.elements.hoveredSelectedItemWithData).toBeDefined()
      expect(items?.elements.hoveredUnselectedItemWithData).toBeDefined()
      expect(items?.elements.hoveredSelectedItemWithNoData).toBeDefined()
      expect(items?.elements.hoveredUnselectedItemWithNoData).toBeDefined()
      for (const [elType, el] of Object.entries(items!.elements)) {
        assertValidDxfColors(el.dxf, `${name}.${elType}`)
      }
    }

    const other1Frame = builtinTableStyle('SlicerStyleOther1')
    expect(other1Frame?.elements.headerRow?.dxf.border?.bottom?.color).toEqual({ rgb: 'FFA6A6A6' })

    for (const name of TIMELINE_NAMES) {
      const frame = builtinTableStyle(name)
      expect(frame, `${name} timeline frame`).toBeDefined()
      expect(frame?.name).toBe(name)
      expect(frame?.elements.wholeTable).toBeDefined()
      expect(frame?.elements.headerRow).toBeDefined()

      const items = builtinSlicerStyle(name)
      expect(items, `${name} timeline items`).toBeDefined()
      expect(items?.name).toBe(name)
      expect(items?.elements.selectedItemWithData).toBeDefined()
      expect(items?.elements.unselectedItemWithData).toBeDefined()
      expect(items?.elements.selectedItemWithData?.dxf.font?.color).toEqual({ rgb: 'FF909090' })
    }

    expect(builtinTableStyle('TimeSlicerStyleOther1')).toBeUndefined()
    expect(builtinSlicerStyle('TimeSlicerStyleOther1')).toBeUndefined()
    expect(builtinTableStyle('TimeSlicerStyleOther2')).toBeUndefined()
    expect(builtinSlicerStyle('TimeSlicerStyleOther2')).toBeUndefined()
    expect(builtinSlicerStyle('TableStyleMedium2')).toBeUndefined()
    expect(builtinTableStyle('NonExistentStyle')).toBeUndefined()
  })

  it('preserves explicit border clear (style: none) on PivotStyleMedium15 headerRow and totalRow', () => {
    const style = builtinTableStyle('PivotStyleMedium15')
    expect(style).toBeDefined()
    const headerRow = style?.elements.headerRow?.dxf
    expect(headerRow).toBeDefined()
    expect(headerRow?.border?.vertical).toEqual({ style: 'none' })
    expect(headerRow?.border?.left).toEqual({ style: 'none' })
    expect(headerRow?.border?.right).toEqual({ style: 'none' })

    const totalRow = style?.elements.totalRow?.dxf
    expect(totalRow).toBeDefined()
    expect(totalRow?.border?.vertical).toEqual({ style: 'none' })
    expect(totalRow?.border?.left).toEqual({ style: 'none' })
    expect(totalRow?.border?.right).toEqual({ style: 'none' })

    // Also check PivotStyleMedium21
    const style21 = builtinTableStyle('PivotStyleMedium21')
    expect(style21?.elements.headerRow?.dxf.border?.left).toEqual({ style: 'none' })
    expect(style21?.elements.headerRow?.dxf.border?.vertical).toEqual({ style: 'none' })
  })

  it('verifies element count and every border side for all 144 styles against committed digest', () => {
    const digestPath = resolve(process.cwd(), 'tests/fixtures/xlsx-pivots/table-styles-digest.json')
    const digest = JSON.parse(readFileSync(digestPath, 'utf8')) as Record<string, {
      elementCount: number
      borders: Record<string, Record<string, string>>
    }>

    const allNames = [...TABLE_NAMES, ...PIVOT_NAMES]
    expect(Object.keys(digest).sort()).toEqual([...allNames].sort())

    for (const name of allNames) {
      const def = builtinTableStyle(name)
      expect(def, `${name} resolves`).toBeDefined()
      const d = digest[name]
      expect(d, `${name} in digest`).toBeDefined()
      expect(Object.keys(def!.elements).length, `${name} elementCount`).toBe(d.elementCount)

      const elements = def!.elements
      for (const [elType, expectedBorders] of Object.entries(d.borders)) {
        const el = elements[elType as keyof typeof elements]
        expect(el, `${name}.${elType} element exists`).toBeDefined()
        const border = el?.dxf.border
        expect(border, `${name}.${elType} has border`).toBeDefined()
        for (const [side, expectedStyle] of Object.entries(expectedBorders)) {
          const sideBorder = border?.[side as keyof typeof border]
          expect(sideBorder?.style, `${name}.${elType}.${side} style`).toBe(expectedStyle)
        }
      }
    }
  })

  it('reports the gzip size of the minified builtins bundle (Ruling S) and enforces <= 12 KB budget', () => {
    const entryPath = resolve(process.cwd(), 'src/xlsx/tableStyles/builtins/index.ts')
    // Bundle with esbuild (Ruling S) in a clean node process to bypass jsdom realm pollution
    const bundled = execFileSync(
      process.execPath,
      [
        '-e',
        `
        const esbuild = require('esbuild');
        esbuild.build({
          entryPoints: [process.argv[1]],
          bundle: true,
          minify: true,
          write: false,
        }).then(r => process.stdout.write(r.outputFiles[0].contents));
        `,
        entryPath,
      ],
      { maxBuffer: 10 * 1024 * 1024 },
    )
    const bundledGzip = gzipSync(bundled).length
    console.log(`[Task 6 Ruling S] Built-in styles minified bundle gzip size: ${bundledGzip} B (budget: 12,288 B)`)
    expect(bundledGzip).toBeLessThanOrEqual(12288)
  })
})
