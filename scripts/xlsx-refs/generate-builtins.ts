/**
 * Generator script for built-in table, pivot, slicer and timeline styles (Task 6).
 *
 * Reads ECMA-376 Part 1 presetTableStyles.xml (from local scratchpad or POI repo)
 * and generates compact representations for:
 *   src/xlsx/tableStyles/builtins/table.ts
 *   src/xlsx/tableStyles/builtins/pivot.ts
 *   src/xlsx/tableStyles/builtins/slicer.ts
 *   src/xlsx/tableStyles/builtins/index.ts
 *
 * Run with: bun scripts/xlsx-refs/generate-builtins.ts
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { XMLParser } from 'fast-xml-parser'
import type { ColorSpec, DxfStyle, TableStyleElementType } from '../../src/xlsx/tableStyles/model'

const SCRATCH_PATH = '/private/tmp/claude-501/-Users-julianshen-prj-officeview/aa69b0c3-1aa3-4cfb-b283-9e7808c48e7e/scratchpad/pv/presetTableStyles.xml'
const POI_URL = 'https://raw.githubusercontent.com/apache/poi/trunk/poi-ooxml/src/main/resources/org/apache/poi/xssf/usermodel/presetTableStyles.xml'

async function ensureXml(): Promise<string> {
  if (existsSync(SCRATCH_PATH)) {
    return readFileSync(SCRATCH_PATH, 'utf8')
  }
  console.log(`Downloading presetTableStyles.xml from ${POI_URL}...`)
  const res = await fetch(POI_URL)
  if (!res.ok) throw new Error(`Failed to download presetTableStyles.xml: ${res.statusText}`)
  const xml = await res.text()
  mkdirSync(dirname(SCRATCH_PATH), { recursive: true })
  writeFileSync(SCRATCH_PATH, xml, 'utf8')
  return xml
}

function parseColor(node: Record<string, unknown> | undefined): ColorSpec | undefined {
  if (!node) return undefined
  const theme = node['@_theme'] !== undefined ? Number(node['@_theme']) : undefined
  const tint = node['@_tint'] !== undefined ? Number(node['@_tint']) : undefined
  if (theme !== undefined) {
    return tint !== undefined ? { theme, tint } : { theme }
  }
  const rgb = node['@_rgb'] as string | undefined
  if (rgb && typeof rgb === 'string') {
    return { rgb: rgb.length === 8 ? rgb.toUpperCase() : `FF${rgb.toUpperCase()}` }
  }
  return undefined
}

function parseDxfNode(dxf: Record<string, unknown> | undefined): DxfStyle {
  if (!dxf) return {}
  const res: DxfStyle = {}
  if (dxf.font) {
    const fNode = dxf.font as Record<string, unknown>
    const font: NonNullable<DxfStyle['font']> = {}
    if (fNode.b !== undefined) font.bold = true
    const c = parseColor(fNode.color as Record<string, unknown> | undefined)
    if (c) font.color = c
    if (Object.keys(font).length) res.font = font
  }
  if (dxf.fill) {
    const fillNode = dxf.fill as Record<string, unknown>
    const pf = fillNode.patternFill as Record<string, unknown> | undefined
    if (pf) {
      const fill: NonNullable<DxfStyle['fill']> = { pattern: (pf['@_patternType'] as string) || 'solid' }
      const fg = parseColor(pf.fgColor as Record<string, unknown> | undefined)
      const bg = parseColor(pf.bgColor as Record<string, unknown> | undefined)
      if (fg) fill.fg = fg
      if (bg) fill.bg = bg
      res.fill = fill
    }
  }
  if (dxf.border) {
    const bNode = dxf.border as Record<string, unknown>
    const border: NonNullable<DxfStyle['border']> = {}
    for (const side of ['left', 'right', 'top', 'bottom', 'horizontal', 'vertical'] as const) {
      const s = bNode[side] as Record<string, unknown> | undefined
      if (s && s['@_style']) {
        const c = parseColor(s.color as Record<string, unknown> | undefined)
        border[side] = c ? { style: s['@_style'] as string, color: c } : { style: s['@_style'] as string }
      }
    }
    if (Object.keys(border).length) res.border = border
  }
  return res
}

export const ELEMENT_TYPES: readonly TableStyleElementType[] = [
  'wholeTable',
  'headerRow',
  'totalRow',
  'firstColumn',
  'lastColumn',
  'firstRowStripe',
  'secondRowStripe',
  'firstColumnStripe',
  'secondColumnStripe',
  'firstHeaderCell',
  'lastHeaderCell',
  'firstTotalCell',
  'lastTotalCell',
  'firstSubtotalColumn',
  'secondSubtotalColumn',
  'thirdSubtotalColumn',
  'firstSubtotalRow',
  'secondSubtotalRow',
  'thirdSubtotalRow',
  'blankRow',
  'firstColumnSubheading',
  'secondColumnSubheading',
  'thirdColumnSubheading',
  'firstRowSubheading',
  'secondRowSubheading',
  'thirdRowSubheading',
  'pageFieldLabels',
  'pageFieldValues',
]

const ELEMENT_TYPE_INDEX = new Map<TableStyleElementType, number>(
  ELEMENT_TYPES.map((t, idx) => [t, idx]),
)

interface ExtractedStyle {
  name: string
  elements: [number, number][] // [typeIndex, dxfIndex]
}

function extractAllStyles(parsed: Record<string, any>): {
  dxfList: DxfStyle[]
  tableStyles: ExtractedStyle[]
  pivotStyles: ExtractedStyle[]
} {
  const dxfList: DxfStyle[] = []
  const dxfIndexMap = new Map<string, number>()
  function getDxfIndex(dxf: DxfStyle): number {
    const k = JSON.stringify(dxf)
    if (dxfIndexMap.has(k)) return dxfIndexMap.get(k)!
    const idx = dxfList.length
    dxfList.push(dxf)
    dxfIndexMap.set(k, idx)
    return idx
  }

  const tableStyles: ExtractedStyle[] = []
  const pivotStyles: ExtractedStyle[] = []

  for (const [key, val] of Object.entries(parsed.presetTableStyles as Record<string, any>)) {
    const ts = val.tableStyles.tableStyle
    const name = ts['@_name'] as string
    const isTable = name.startsWith('TableStyle')
    const isPivot = name.startsWith('PivotStyle')
    if (!isTable && !isPivot) continue

    const rawDxfs = Array.isArray(val.dxfs.dxf) ? val.dxfs.dxf : [val.dxfs.dxf]
    const parsedDxfs = rawDxfs.map(parseDxfNode)
    const elements: [number, number][] = []
    const rawElements = Array.isArray(ts.tableStyleElement) ? ts.tableStyleElement : [ts.tableStyleElement]
    for (const el of rawElements) {
      const type = el['@_type'] as TableStyleElementType
      const typeIdx = ELEMENT_TYPE_INDEX.get(type)
      if (typeIdx === undefined) throw new Error(`Unknown element type: ${type}`)
      // TableStyle dxfId is 1-based, PivotStyle dxfId is 0-based
      const rawId = parseInt(el['@_dxfId'], 10)
      const id = isTable ? rawId - 1 : rawId
      const dxf = parsedDxfs[id]
      if (dxf) elements.push([typeIdx, getDxfIndex(dxf)])
    }

    if (isTable) {
      tableStyles.push({ name, elements })
    } else {
      pivotStyles.push({ name, elements })
    }
  }

  return { dxfList, tableStyles, pivotStyles }
}

function generateTableTs(dxfList: DxfStyle[], tableStyles: ExtractedStyle[]): string {
  const mapObj: Record<string, [number, number][]> = {}
  for (const s of tableStyles) {
    mapObj[s.name] = s.elements
  }

  return `/**
 * Built-in Excel table styles (TableStyleLight1-21, TableStyleMedium1-28, TableStyleDark1-11).
 * Transcribed from ECMA-376 Part 1 presetTableStyles.xml.
 */
import type { DxfStyle, TableStyleDef, TableStyleElementType } from '../model'

export const ELEMENT_TYPES: readonly TableStyleElementType[] = ${JSON.stringify(ELEMENT_TYPES)}

export const DXFS: readonly DxfStyle[] = ${JSON.stringify(dxfList)}

const TABLES: Readonly<Record<string, readonly (readonly [number, number])[]>> = ${JSON.stringify(mapObj)}

export function getTableStyle(name: string): TableStyleDef | undefined {
  const entries = TABLES[name]
  if (!entries) return undefined
  const elements: Partial<Record<TableStyleElementType, { dxf: DxfStyle }>> = {}
  for (const [tIdx, dIdx] of entries) {
    elements[ELEMENT_TYPES[tIdx]!] = { dxf: DXFS[dIdx]! }
  }
  return { name, pivot: false, table: true, elements }
}
`
}

function generatePivotTs(pivotStyles: ExtractedStyle[]): string {
  const mapObj: Record<string, [number, number][]> = {}
  for (const s of pivotStyles) {
    mapObj[s.name] = s.elements
  }

  return `/**
 * Built-in Excel pivot styles (PivotStyleLight1-28, PivotStyleMedium1-28, PivotStyleDark1-28).
 * Transcribed from ECMA-376 Part 1 presetTableStyles.xml.
 */
import type { DxfStyle, TableStyleDef, TableStyleElementType } from '../model'
import { DXFS, ELEMENT_TYPES } from './table'

const PIVOTS: Readonly<Record<string, readonly (readonly [number, number])[]>> = ${JSON.stringify(mapObj)}

export function getPivotStyle(name: string): TableStyleDef | undefined {
  const entries = PIVOTS[name]
  if (!entries) return undefined
  const elements: Partial<Record<TableStyleElementType, { dxf: DxfStyle }>> = {}
  for (const [tIdx, dIdx] of entries) {
    elements[ELEMENT_TYPES[tIdx]!] = { dxf: DXFS[dIdx]! }
  }
  return { name, pivot: true, table: false, elements }
}
`
}

function generateSlicerTs(): string {
  return `/**
 * Built-in Excel slicer and timeline styles.
 * Derived from Task 3 reference captures (tests/fixtures/xlsx-pivots/refs/).
 * Annotated with [measure] tags for reference IDs and measured colors.
 */
import type { ColorSpec, DxfStyle, SlicerStyleDef, TableStyleDef } from '../model'

// Common border and font helpers
const thinBorder = (color: ColorSpec): NonNullable<DxfStyle['border']> => ({
  left: { style: 'thin', color },
  right: { style: 'thin', color },
  top: { style: 'thin', color },
  bottom: { style: 'thin', color },
})

const solidFill = (bg: ColorSpec): NonNullable<DxfStyle['fill']> => ({
  pattern: 'solid',
  bg,
})

// [measure] refs/slicer-light1.png through refs/slicer-light6.png:
// Frame border is 1 px in accent colour (themes 4 to 9).
// Header caption on white background with 1 px accent separator at bottom.
// Selected item has solid accent tint 0.6 fill with #999999 border.
// Unselected item has white fill with #CCCCCC border.
// [measure] refs/slicer-no-data.png:
// Item with no data has white fill with #E0E0E0 border and #828282 text.

function makeSlicerLightStyle(index: number): { tableDef: TableStyleDef; slicerDef: SlicerStyleDef } {
  const theme = 3 + index
  const name = \`SlicerStyleLight\${index}\`
  const tableDef: TableStyleDef = {
    name,
    pivot: false,
    table: false,
    elements: {
      // [measure] refs/slicer-light\${index}.png: frame border in style accent
      wholeTable: { dxf: { border: thinBorder({ theme }) } },
      // [measure] refs/slicer-light\${index}.png: header on white with accent separator
      headerRow: {
        dxf: {
          font: { bold: true, color: { theme: 1 } },
          fill: solidFill({ rgb: 'FFFFFFFF' }),
          border: { bottom: { style: 'thin', color: { theme } } },
        },
      },
    },
  }

  const slicerDef: SlicerStyleDef = {
    name,
    elements: {
      // [measure] refs/slicer-light\${index}.png: selected item fill theme tint 0.6, border #999999
      selectedItemWithData: {
        dxf: {
          fill: solidFill({ theme, tint: 0.6 }),
          border: thinBorder({ rgb: 'FF999999' }),
          font: { color: { theme: 1 } },
        },
      },
      // [measure] refs/slicer-light\${index}.png: unselected item fill white, border #CCCCCC
      unselectedItemWithData: {
        dxf: {
          fill: solidFill({ rgb: 'FFFFFFFF' }),
          border: thinBorder({ rgb: 'FFCCCCCC' }),
          font: { color: { theme: 1 } },
        },
      },
      // [measure] refs/slicer-no-data.png: selected no-data item pale accent tint 0.8, border #E0E0E0, text #828282
      selectedItemWithNoData: {
        dxf: {
          fill: solidFill({ theme, tint: 0.8 }),
          border: thinBorder({ rgb: 'FFE0E0E0' }),
          font: { color: { rgb: 'FF828282' } },
        },
      },
      // [measure] refs/slicer-no-data.png: unselected no-data item fill white, border #E0E0E0, text #828282
      unselectedItemWithNoData: {
        dxf: {
          fill: solidFill({ rgb: 'FFFFFFFF' }),
          border: thinBorder({ rgb: 'FFE0E0E0' }),
          font: { color: { rgb: 'FF828282' } },
        },
      },
      hoveredSelectedItemWithData: {
        dxf: {
          fill: solidFill({ theme, tint: 0.5 }),
          border: thinBorder({ rgb: 'FF999999' }),
          font: { color: { theme: 1 } },
        },
      },
      hoveredUnselectedItemWithData: {
        dxf: {
          fill: solidFill({ theme, tint: 0.85 }),
          border: thinBorder({ rgb: 'FFCCCCCC' }),
          font: { color: { theme: 1 } },
        },
      },
      hoveredSelectedItemWithNoData: {
        dxf: {
          fill: solidFill({ theme, tint: 0.85 }),
          border: thinBorder({ rgb: 'FFE0E0E0' }),
          font: { color: { rgb: 'FF828282' } },
        },
      },
      hoveredUnselectedItemWithNoData: {
        dxf: {
          fill: solidFill({ rgb: 'FFF5F5F5' }),
          border: thinBorder({ rgb: 'FFE0E0E0' }),
          font: { color: { rgb: 'FF828282' } },
        },
      },
    },
  }
  return { tableDef, slicerDef }
}

// [measure] refs/slicer-dark1.png through refs/slicer-dark6.png:
// Frame border is 1 px in accent colour (themes 4 to 9).
// Selected item has solid accent fill and border with bold white text.
// Unselected item has silver fill #C0C0C0 and border #C0C0C0 with dark text.

function makeSlicerDarkStyle(index: number): { tableDef: TableStyleDef; slicerDef: SlicerStyleDef } {
  const theme = 3 + index
  const name = \`SlicerStyleDark\${index}\`
  const tableDef: TableStyleDef = {
    name,
    pivot: false,
    table: false,
    elements: {
      // [measure] refs/slicer-dark\${index}.png: frame border in style accent
      wholeTable: { dxf: { border: thinBorder({ theme }) } },
      // [measure] refs/slicer-dark\${index}.png: header on white with accent separator
      headerRow: {
        dxf: {
          font: { bold: true, color: { theme: 1 } },
          fill: solidFill({ rgb: 'FFFFFFFF' }),
          border: { bottom: { style: 'thin', color: { theme } } },
        },
      },
    },
  }

  const slicerDef: SlicerStyleDef = {
    name,
    elements: {
      // [measure] refs/slicer-dark\${index}.png: selected item solid accent fill, matching border, bold white text
      selectedItemWithData: {
        dxf: {
          fill: solidFill({ theme }),
          border: thinBorder({ theme }),
          font: { bold: true, color: { theme: 0 } },
        },
      },
      // [measure] refs/slicer-dark\${index}.png: unselected item silver #C0C0C0 fill and border, dark text
      unselectedItemWithData: {
        dxf: {
          fill: solidFill({ rgb: 'FFC0C0C0' }),
          border: thinBorder({ rgb: 'FFC0C0C0' }),
          font: { color: { theme: 1 } },
        },
      },
      selectedItemWithNoData: {
        dxf: {
          fill: solidFill({ theme, tint: 0.4 }),
          border: thinBorder({ theme, tint: 0.4 }),
          font: { color: { rgb: 'FFE0E0E0' } },
        },
      },
      unselectedItemWithNoData: {
        dxf: {
          fill: solidFill({ rgb: 'FFD4D4D4' }),
          border: thinBorder({ rgb: 'FFD4D4D4' }),
          font: { color: { rgb: 'FF828282' } },
        },
      },
      hoveredSelectedItemWithData: {
        dxf: {
          fill: solidFill({ theme, tint: -0.15 }),
          border: thinBorder({ theme, tint: -0.15 }),
          font: { bold: true, color: { theme: 0 } },
        },
      },
      hoveredUnselectedItemWithData: {
        dxf: {
          fill: solidFill({ rgb: 'FFD0D0D0' }),
          border: thinBorder({ rgb: 'FFD0D0D0' }),
          font: { color: { theme: 1 } },
        },
      },
      hoveredSelectedItemWithNoData: {
        dxf: {
          fill: solidFill({ theme, tint: 0.3 }),
          border: thinBorder({ theme, tint: 0.3 }),
          font: { color: { rgb: 'FFE0E0E0' } },
        },
      },
      hoveredUnselectedItemWithNoData: {
        dxf: {
          fill: solidFill({ rgb: 'FFE0E0E0' }),
          border: thinBorder({ rgb: 'FFE0E0E0' }),
          font: { color: { rgb: 'FF828282' } },
        },
      },
    },
  }
  return { tableDef, slicerDef }
}

// [measure] refs/slicer-other1.png:
// Frame border #808080.
// Selected item has fill #BFBFBF, border #999999.
// Unselected item has fill white, border #CCCCCC.
const slicerOther1Table: TableStyleDef = {
  name: 'SlicerStyleOther1',
  pivot: false,
  table: false,
  elements: {
    wholeTable: { dxf: { border: thinBorder({ rgb: 'FF808080' }) } },
    headerRow: {
      dxf: {
        font: { bold: true, color: { theme: 1 } },
        fill: solidFill({ rgb: 'FFFFFFFF' }),
        border: { bottom: { style: 'thin', color: { rgb: 'FF808080' } } },
      },
    },
  },
}

const slicerOther1Slicer: SlicerStyleDef = {
  name: 'SlicerStyleOther1',
  elements: {
    selectedItemWithData: {
      dxf: {
        fill: solidFill({ rgb: 'FFBFBFBF' }),
        border: thinBorder({ rgb: 'FF999999' }),
        font: { color: { theme: 1 } },
      },
    },
    unselectedItemWithData: {
      dxf: {
        fill: solidFill({ rgb: 'FFFFFFFF' }),
        border: thinBorder({ rgb: 'FFCCCCCC' }),
        font: { color: { theme: 1 } },
      },
    },
    selectedItemWithNoData: {
      dxf: {
        fill: solidFill({ rgb: 'FFD4D4D4' }),
        border: thinBorder({ rgb: 'FFE0E0E0' }),
        font: { color: { rgb: 'FF828282' } },
      },
    },
    unselectedItemWithNoData: {
      dxf: {
        fill: solidFill({ rgb: 'FFFFFFFF' }),
        border: thinBorder({ rgb: 'FFE0E0E0' }),
        font: { color: { rgb: 'FF828282' } },
      },
    },
  },
}

// [measure] refs/slicer-other2.png:
// Frame border Accent 1 #4F81BD.
// Selected item has fill #C2D8F1, border #999999.
// Unselected item has fill #F4F5F7, border #CCCCCC.
const slicerOther2Table: TableStyleDef = {
  name: 'SlicerStyleOther2',
  pivot: false,
  table: false,
  elements: {
    wholeTable: { dxf: { border: thinBorder({ theme: 4 }) } },
    headerRow: {
      dxf: {
        font: { bold: true, color: { theme: 1 } },
        fill: solidFill({ rgb: 'FFFFFFFF' }),
        border: { bottom: { style: 'thin', color: { theme: 4 } } },
      },
    },
  },
}

const slicerOther2Slicer: SlicerStyleDef = {
  name: 'SlicerStyleOther2',
  elements: {
    selectedItemWithData: {
      dxf: {
        fill: solidFill({ rgb: 'FFC2D8F1' }),
        border: thinBorder({ rgb: 'FF999999' }),
        font: { color: { theme: 1 } },
      },
    },
    unselectedItemWithData: {
      dxf: {
        fill: solidFill({ rgb: 'FFF4F5F7' }),
        border: thinBorder({ rgb: 'FFCCCCCC' }),
        font: { color: { theme: 1 } },
      },
    },
    selectedItemWithNoData: {
      dxf: {
        fill: solidFill({ rgb: 'FFE2EAF5' }),
        border: thinBorder({ rgb: 'FFE0E0E0' }),
        font: { color: { rgb: 'FF828282' } },
      },
    },
    unselectedItemWithNoData: {
      dxf: {
        fill: solidFill({ rgb: 'FFFFFFFF' }),
        border: thinBorder({ rgb: 'FFE0E0E0' }),
        font: { color: { rgb: 'FF828282' } },
      },
    },
  },
}

// [measure] refs/timeline-light1-months.png through refs/timeline-light6-months.png:
// Timeline frame border in accent colour (themes 4 to 9).
// Time bar light accent tint fill #8FB1D6, dark shade border.
function makeTimelineLightStyle(index: number): { tableDef: TableStyleDef; slicerDef: SlicerStyleDef } {
  const theme = 3 + index
  const name = \`TimeSlicerStyleLight\${index}\`
  const tableDef: TableStyleDef = {
    name,
    pivot: false,
    table: false,
    elements: {
      // [measure] refs/timeline-light\${index}-months.png: frame border in style accent
      wholeTable: { dxf: { border: thinBorder({ theme }) } },
      // [measure] refs/timeline-light\${index}-months.png: header separator in style accent
      headerRow: {
        dxf: {
          font: { bold: true, color: { theme: 1 } },
          fill: solidFill({ rgb: 'FFFFFFFF' }),
          border: { bottom: { style: 'thin', color: { theme } } },
        },
      },
    },
  }

  const slicerDef: SlicerStyleDef = {
    name,
    elements: {
      // [measure] refs/timeline-light\${index}-months.png: selected period bar light accent fill
      selectedItemWithData: {
        dxf: {
          fill: solidFill({ theme, tint: 0.4 }),
          border: thinBorder({ theme }),
          font: { color: { theme: 1 } },
        },
      },
      unselectedItemWithData: {
        dxf: {
          fill: solidFill({ rgb: 'FFFFFFFF' }),
          border: thinBorder({ rgb: 'FFE0E0E0' }),
          font: { color: { theme: 1 } },
        },
      },
    },
  }
  return { tableDef, slicerDef }
}

// [measure] refs/timeline-dark1-months.png through refs/timeline-dark6-months.png:
// Timeline frame border in accent colour (themes 4 to 9).
// Time bar dark accent shade fill #39699A.
function makeTimelineDarkStyle(index: number): { tableDef: TableStyleDef; slicerDef: SlicerStyleDef } {
  const theme = 3 + index
  const name = \`TimeSlicerStyleDark\${index}\`
  const tableDef: TableStyleDef = {
    name,
    pivot: false,
    table: false,
    elements: {
      // [measure] refs/timeline-dark\${index}-months.png: frame border in style accent
      wholeTable: { dxf: { border: thinBorder({ theme }) } },
      // [measure] refs/timeline-dark\${index}-months.png: header separator in style accent
      headerRow: {
        dxf: {
          font: { bold: true, color: { theme: 1 } },
          fill: solidFill({ rgb: 'FFFFFFFF' }),
          border: { bottom: { style: 'thin', color: { theme } } },
        },
      },
    },
  }

  const slicerDef: SlicerStyleDef = {
    name,
    elements: {
      // [measure] refs/timeline-dark\${index}-months.png: selected period bar dark accent shade fill
      selectedItemWithData: {
        dxf: {
          fill: solidFill({ theme, tint: -0.2 }),
          border: thinBorder({ theme }),
          font: { color: { theme: 0 } },
        },
      },
      unselectedItemWithData: {
        dxf: {
          fill: solidFill({ rgb: 'FFFFFFFF' }),
          border: thinBorder({ rgb: 'FFE0E0E0' }),
          font: { color: { theme: 1 } },
        },
      },
    },
  }
  return { tableDef, slicerDef }
}

// Build maps
const SLICER_TABLE_STYLES = new Map<string, TableStyleDef>()
const SLICER_STYLES = new Map<string, SlicerStyleDef>()

for (let i = 1; i <= 6; i++) {
  const light = makeSlicerLightStyle(i)
  SLICER_TABLE_STYLES.set(light.tableDef.name, light.tableDef)
  SLICER_STYLES.set(light.slicerDef.name, light.slicerDef)

  const dark = makeSlicerDarkStyle(i)
  SLICER_TABLE_STYLES.set(dark.tableDef.name, dark.tableDef)
  SLICER_STYLES.set(dark.slicerDef.name, dark.slicerDef)

  const tlLight = makeTimelineLightStyle(i)
  SLICER_TABLE_STYLES.set(tlLight.tableDef.name, tlLight.tableDef)
  SLICER_STYLES.set(tlLight.slicerDef.name, tlLight.slicerDef)

  const tlDark = makeTimelineDarkStyle(i)
  SLICER_TABLE_STYLES.set(tlDark.tableDef.name, tlDark.tableDef)
  SLICER_STYLES.set(tlDark.slicerDef.name, tlDark.slicerDef)
}

SLICER_TABLE_STYLES.set(slicerOther1Table.name, slicerOther1Table)
SLICER_STYLES.set(slicerOther1Slicer.name, slicerOther1Slicer)
SLICER_TABLE_STYLES.set(slicerOther2Table.name, slicerOther2Table)
SLICER_STYLES.set(slicerOther2Slicer.name, slicerOther2Slicer)

export function getSlicerTableStyle(name: string): TableStyleDef | undefined {
  return SLICER_TABLE_STYLES.get(name)
}

export function getSlicerStyle(name: string): SlicerStyleDef | undefined {
  return SLICER_STYLES.get(name)
}
`
}

function generateIndexTs(): string {
  return `/**
 * Built-in Excel table, pivot, slicer and timeline style provider.
 * Lazy chunk: eager code may only import type from here.
 */
import { getTableStyle } from './table'
import { getPivotStyle } from './pivot'
import { getSlicerTableStyle, getSlicerStyle } from './slicer'
import type { SlicerStyleDef, TableStyleDef } from '../model'

/**
 * Returns the TableStyleDef for any built-in table style, pivot style,
 * or slicer/timeline frame style. Returns undefined for unknown names.
 */
export function builtinTableStyle(name: string): TableStyleDef | undefined {
  if (name.startsWith('TableStyle')) return getTableStyle(name)
  if (name.startsWith('PivotStyle')) return getPivotStyle(name)
  if (name.startsWith('SlicerStyle') || name.startsWith('TimeSlicerStyle')) return getSlicerTableStyle(name)
  return undefined
}

/**
 * Returns the SlicerStyleDef (item states) for any built-in slicer or timeline style.
 * Returns undefined for table styles, pivot styles, or unknown names.
 */
export function builtinSlicerStyle(name: string): SlicerStyleDef | undefined {
  if (name.startsWith('SlicerStyle') || name.startsWith('TimeSlicerStyle')) return getSlicerStyle(name)
  return undefined
}
`
}

export async function generateBuiltins(): Promise<void> {
  const xml = await ensureXml()
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })
  const parsed = parser.parse(xml)

  const outDir = resolve(import.meta.dir, '../../src/xlsx/tableStyles/builtins')
  mkdirSync(outDir, { recursive: true })

  const { dxfList, tableStyles, pivotStyles } = extractAllStyles(parsed)
  console.log(`Extracted ${dxfList.length} unique DXFs, ${tableStyles.length} table styles, ${pivotStyles.length} pivot styles.`)

  const tableTs = generateTableTs(dxfList, tableStyles)
  writeFileSync(resolve(outDir, 'table.ts'), tableTs, 'utf8')
  console.log(`wrote ${resolve(outDir, 'table.ts')} (${tableTs.length} bytes)`)

  const pivotTs = generatePivotTs(pivotStyles)
  writeFileSync(resolve(outDir, 'pivot.ts'), pivotTs, 'utf8')
  console.log(`wrote ${resolve(outDir, 'pivot.ts')} (${pivotTs.length} bytes)`)

  const slicerTs = generateSlicerTs()
  writeFileSync(resolve(outDir, 'slicer.ts'), slicerTs, 'utf8')
  console.log(`wrote ${resolve(outDir, 'slicer.ts')} (${slicerTs.length} bytes)`)

  const indexTs = generateIndexTs()
  writeFileSync(resolve(outDir, 'index.ts'), indexTs, 'utf8')
  console.log(`wrote ${resolve(outDir, 'index.ts')} (${indexTs.length} bytes)`)
}

if (import.meta.main) {
  await generateBuiltins()
  console.log('Done generating built-in styles.')
}
