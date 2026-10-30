/**
 * Hand-written style inputs derived from the Excel-made pivot seed (pivots plan Task 3). Excel opens each one,
 * refreshes and re-saves it (build-pivots.ts), so the committed references are Excel's own output.
 *
 * - pivot-custom-style-input.xlsx: custom pivot style `OfficeViewPivot` (the `pivot-custom-style` case).
 * - pivot-precedence-input.xlsx: custom style `OfficeViewPrecedence`, one distinct fill per element with every
 *   style-info flag on, so the colour at each overlap shows which element wins (spec §6.5 precedence).
 * - pivot-cell-override-input.xlsx: three pivot cells get their own fill through `s=` with no pivot `formats`
 *   entry, to measure how a cell's own format layers over the pivot style (spec §6.5 **[measure]**).
 *
 *   bun scripts/xlsx-refs/style-inputs.ts
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import JSZip from 'jszip'

const RECIPES = resolve(import.meta.dir, '../../tests/fixtures/xlsx-pivots/recipes')
const EMPTY_STYLES = '<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>'
/** The seed's cellXfs count; the override xf is appended at this index. */
const SEED_XF_COUNT = 5
const SEED_STYLE_INFO = /<pivotTableStyleInfo [^>]*\/>/

const fill = (rgb: string) => `<fill><patternFill patternType="solid"><bgColor rgb="FF${rgb}"/></patternFill></fill>`
const side = (n: string) => `<${n} style="thin"><color rgb="FF7F7F7F"/></${n}>`

function styleSheet(name: string, elements: readonly (readonly [string, string])[]): string {
  return `<dxfs count="${elements.length}">${elements.map(([, dxf]) => `<dxf>${dxf}</dxf>`).join('')}</dxfs>`
    + `<tableStyles count="1" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"><tableStyle name="${name}" table="0" count="${elements.length}">`
    + elements.map(([type], i) => `<tableStyleElement type="${type}" dxfId="${i}"/>`).join('') + '</tableStyle></tableStyles>'
}

/** Pivot table style element types (ECMA-376 §18.18.82) that the seed pivot's layout can show. */
export const PRECEDENCE_ELEMENTS = [
  'wholeTable', 'headerRow', 'totalRow', 'firstColumn', 'lastColumn', 'firstRowStripe', 'secondRowStripe',
  'firstColumnStripe', 'secondColumnStripe', 'firstHeaderCell', 'firstSubtotalColumn', 'secondSubtotalColumn',
  'firstSubtotalRow', 'secondSubtotalRow', 'blankRow', 'firstColumnSubheading', 'secondColumnSubheading',
  'firstRowSubheading', 'secondRowSubheading',
] as const

/** Evenly spaced hues, so every element's fill is distinguishable in a capture. */
export function precedenceColors(): Record<string, string> {
  const hex = (v: number) => Math.round(v * 255).toString(16).padStart(2, '0').toUpperCase()
  return Object.fromEntries(PRECEDENCE_ELEMENTS.map((type, i) => {
    const h = i / PRECEDENCE_ELEMENTS.length, s = 0.55, v = 0.95
    const k = (n: number) => (n + h * 6) % 6
    const f = (n: number) => v - v * s * Math.max(0, Math.min(k(n), 4 - k(n), 1))
    return [type, hex(f(5)) + hex(f(3)) + hex(f(1))]
  }))
}

async function derive(out: string, edit: (path: string, xml: string) => string): Promise<void> {
  const zip = await JSZip.loadAsync(readFileSync(join(RECIPES, 'pivot-seed.xlsx')))
  for (const path of Object.keys(zip.files)) {
    if (!path.endsWith('.xml')) continue
    const xml = await zip.file(path)!.async('text')
    const next = edit(path, xml)
    if (next !== xml) zip.file(path, next)
  }
  writeFileSync(join(RECIPES, out), await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}

const swap = (xml: string, from: string | RegExp, to: string, what: string): string => {
  const next = xml.replace(from, to)
  if (next === xml) throw new Error(`seed changed shape: ${what}`)
  return next
}

export async function writeStyleInputs(): Promise<void> {
  await derive('pivot-custom-style-input.xlsx', (path, xml) => {
    if (path === 'xl/styles.xml') return swap(xml, EMPTY_STYLES, styleSheet('OfficeViewPivot', [
      ['wholeTable', `<border>${['left', 'right', 'top', 'bottom'].map(side).join('')}</border>`],
      ['headerRow', `<font><b/><color rgb="FFFFFFFF"/></font>${fill('1F4E78')}`],
      ['totalRow', `<font><b/></font>${fill('D9D9D9')}`],
      ['firstRowSubheading', `<font><b/><i/></font>${fill('FCE4D6')}`],
      ['firstColumn', '<font><color rgb="FFC00000"/></font>'],
    ]), 'styles')
    if (path.startsWith('xl/pivotTables/')) return swap(xml, 'name="PivotStyleLight16"', 'name="OfficeViewPivot"', 'pivot style')
    return xml
  })

  const colors = precedenceColors()
  await derive('pivot-precedence-input.xlsx', (path, xml) => {
    if (path === 'xl/styles.xml') return swap(xml, EMPTY_STYLES, styleSheet('OfficeViewPrecedence', PRECEDENCE_ELEMENTS.map(t => [t, fill(colors[t])] as const)), 'styles')
    if (path.startsWith('xl/pivotTables/')) return swap(xml, SEED_STYLE_INFO, '<pivotTableStyleInfo name="OfficeViewPrecedence" showRowHeaders="1" showColHeaders="1" showRowStripes="1" showColStripes="1" showLastColumn="1"/>', 'style info')
    return xml
  })
  writeFileSync(join(RECIPES, 'pivot-precedence-colors.json'), JSON.stringify(colors, null, 2) + '\n')

  // Green fill + bold red font on a value cell (C7), a row label (A6, "East") and a grand-total cell (B22).
  await derive('pivot-cell-override-input.xlsx', (path, xml) => {
    if (path === 'xl/styles.xml') {
      let next = swap(xml, /<fills count="2">(.*?)<\/fills>/, '<fills count="3">$1<fill><patternFill patternType="solid"><fgColor rgb="FF92D050"/><bgColor indexed="64"/></patternFill></fill></fills>', 'fills')
      const fontCount = Number(/<fonts count="(\d+)">/.exec(next)![1])
      next = next.replace(/<fonts count="\d+">(.*?)<\/fonts>/, `<fonts count="${fontCount + 1}">$1<font><b/><sz val="11"/><color rgb="FFC00000"/><name val="Calibri"/><family val="2"/></font></fonts>`)
      const xfCount = Number(/<cellXfs count="(\d+)">/.exec(next)![1])
      if (xfCount !== SEED_XF_COUNT) throw new Error(`seed changed shape: ${xfCount} cellXfs`)
      next = next.replace(/<cellXfs count="\d+">(.*?)<\/cellXfs>/, `<cellXfs count="${xfCount + 1}">$1<xf numFmtId="0" fontId="${fontCount}" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"><alignment vertical="center"/></xf></cellXfs>`)
      return next
    }
    if (path === 'xl/worksheets/sheet1.xml') {
      let next = xml
      for (const ref of ['A6', 'C7', 'B22']) next = swap(next, new RegExp(`<c r="${ref}" s="\\d+"`), `<c r="${ref}" s="${SEED_XF_COUNT}"`, ref)
      return next
    }
    return xml
  })
}

if (import.meta.main) {
  await writeStyleInputs()
  console.log(`wrote style inputs to ${RECIPES}`)
}
