/**
 * Generator script for built-in table, pivot, slicer and timeline styles (Task 6).
 *
 * Source for table/pivot styles: ECMA-376 Part 1 presetTableStyles.xml as shipped by Apache POI.
 * Note: in presetTableStyles.xml, TableStyle dxfId is 1-based, while PivotStyle dxfId is 0-based.
 *
 * Slicer/timeline styles are not in ECMA-376; they are derived from Task 3 reference captures
 * in tests/fixtures/xlsx-pivots/refs/.
 *
 * Generates:
 *   src/xlsx/tableStyles/builtins/table.ts
 *   src/xlsx/tableStyles/builtins/pivot.ts
 *   src/xlsx/tableStyles/builtins/slicer.ts
 *   src/xlsx/tableStyles/builtins/index.ts
 *   tests/fixtures/xlsx-pivots/table-styles-digest.json
 *
 * Run with: bun scripts/xlsx-refs/generate-builtins.ts [path-to-presetTableStyles.xml]
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { XMLParser } from 'fast-xml-parser'
import type { ColorSpec, DxfStyle, TableStyleElementType } from '../../src/xlsx/tableStyles/model'

const DEFAULT_CACHE_PATH = resolve(process.cwd(), '.tmp/presetTableStyles.xml')
const SCRATCH_PATH = '/private/tmp/claude-501/-Users-julianshen-prj-officeview/aa69b0c3-1aa3-4cfb-b283-9e7808c48e7e/scratchpad/pv/presetTableStyles.xml'
const POI_URL = 'https://raw.githubusercontent.com/apache/poi/trunk/poi-ooxml/src/main/resources/org/apache/poi/xssf/usermodel/presetTableStyles.xml'

async function ensureXml(cliPath?: string): Promise<string> {
  if (cliPath && existsSync(cliPath)) {
    return readFileSync(cliPath, 'utf8')
  }
  const envPath = process.env.PRESET_TABLE_STYLES_XML
  if (envPath && existsSync(envPath)) {
    return readFileSync(envPath, 'utf8')
  }
  if (existsSync(SCRATCH_PATH)) {
    return readFileSync(SCRATCH_PATH, 'utf8')
  }
  if (existsSync(DEFAULT_CACHE_PATH)) {
    return readFileSync(DEFAULT_CACHE_PATH, 'utf8')
  }
  console.log(`Downloading presetTableStyles.xml from ${POI_URL}...`)
  const res = await fetch(POI_URL)
  if (!res.ok) throw new Error(`Failed to download presetTableStyles.xml: ${res.statusText}`)
  const xml = await res.text()
  mkdirSync(dirname(DEFAULT_CACHE_PATH), { recursive: true })
  writeFileSync(DEFAULT_CACHE_PATH, xml, 'utf8')
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
      if (bNode[side] !== undefined) {
        const s = bNode[side] as Record<string, unknown> | string
        if (typeof s === 'object' && s !== null && s['@_style']) {
          const c = parseColor(s.color as Record<string, unknown> | undefined)
          border[side] = c ? { style: s['@_style'] as string, color: c } : { style: s['@_style'] as string }
        } else {
          // Empty element (e.g. <left/>) inside dxf border is an explicit clear
          border[side] = { style: 'none' }
        }
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
  rawElementCount: number
  borders: Partial<Record<TableStyleElementType, Partial<Record<string, string>>>>
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
    const borders: Partial<Record<TableStyleElementType, Partial<Record<string, string>>>> = {}

    for (const el of rawElements) {
      const type = el['@_type'] as TableStyleElementType
      const typeIdx = ELEMENT_TYPE_INDEX.get(type)
      if (typeIdx === undefined) throw new Error(`Unknown element type: ${type}`)
      // TableStyle dxfId is 1-based, PivotStyle dxfId is 0-based
      const rawId = parseInt(el['@_dxfId'], 10)
      const id = isTable ? rawId - 1 : rawId
      const dxf = parsedDxfs[id]
      if (dxf) {
        elements.push([typeIdx, getDxfIndex(dxf)])
        if (dxf.border) {
          const sideMap: Partial<Record<string, string>> = {}
          for (const [side, b] of Object.entries(dxf.border)) {
            if (b?.style) sideMap[side] = b.style
          }
          if (Object.keys(sideMap).length > 0) {
            borders[type] = sideMap
          }
        }
      }
    }

    const item: ExtractedStyle = {
      name,
      elements,
      rawElementCount: rawElements.length,
      borders,
    }

    if (isTable) {
      tableStyles.push(item)
    } else {
      pivotStyles.push(item)
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
 * Transcribed from ECMA-376 presetTableStyles.xml as shipped by Apache POI.
 * Note: TableStyle dxfId is 1-based in presetTableStyles.xml.
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
 * Transcribed from ECMA-376 presetTableStyles.xml as shipped by Apache POI.
 * Note: PivotStyle dxfId is 0-based in presetTableStyles.xml.
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
 *
 * Annotations:
 *   [measure] <ref id> - value measured directly from the referenced capture
 *   [approx] - approximation of complex appearance (e.g. flat fill for vertical gradient)
 *   [estimate] - state not visible in still captures, estimated based on theme convention
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
      // [estimate: not visible in still captures] selected no-data pale accent tint 0.8
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
      // [estimate: hover state not measurable from still captures]
      hoveredSelectedItemWithData: {
        dxf: {
          fill: solidFill({ theme, tint: 0.5 }),
          border: thinBorder({ rgb: 'FF999999' }),
          font: { color: { theme: 1 } },
        },
      },
      // [estimate: hover state not measurable from still captures]
      hoveredUnselectedItemWithData: {
        dxf: {
          fill: solidFill({ theme, tint: 0.85 }),
          border: thinBorder({ rgb: 'FFCCCCCC' }),
          font: { color: { theme: 1 } },
        },
      },
      // [estimate: hover state not measurable from still captures]
      hoveredSelectedItemWithNoData: {
        dxf: {
          fill: solidFill({ theme, tint: 0.85 }),
          border: thinBorder({ rgb: 'FFE0E0E0' }),
          font: { color: { rgb: 'FF828282' } },
        },
      },
      // [estimate: hover state not measurable from still captures]
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
      // [estimate: not visible in still captures]
      selectedItemWithNoData: {
        dxf: {
          fill: solidFill({ theme, tint: 0.4 }),
          border: thinBorder({ theme, tint: 0.4 }),
          font: { color: { rgb: 'FFE0E0E0' } },
        },
      },
      // [estimate: not visible in still captures]
      unselectedItemWithNoData: {
        dxf: {
          fill: solidFill({ rgb: 'FFD4D4D4' }),
          border: thinBorder({ rgb: 'FFD4D4D4' }),
          font: { color: { rgb: 'FF828282' } },
        },
      },
      // [estimate: hover state not measurable from still captures]
      hoveredSelectedItemWithData: {
        dxf: {
          fill: solidFill({ theme, tint: -0.15 }),
          border: thinBorder({ theme, tint: -0.15 }),
          font: { bold: true, color: { theme: 0 } },
        },
      },
      // [estimate: hover state not measurable from still captures]
      hoveredUnselectedItemWithData: {
        dxf: {
          fill: solidFill({ rgb: 'FFD0D0D0' }),
          border: thinBorder({ rgb: 'FFD0D0D0' }),
          font: { color: { theme: 1 } },
        },
      },
      // [estimate: hover state not measurable from still captures]
      hoveredSelectedItemWithNoData: {
        dxf: {
          fill: solidFill({ theme, tint: 0.3 }),
          border: thinBorder({ theme, tint: 0.3 }),
          font: { color: { rgb: 'FFE0E0E0' } },
        },
      },
      // [estimate: hover state not measurable from still captures]
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

// [measure] refs/slicer-other1.png: frame border #808080, header separator #A6A6A6
const slicerOther1Table: TableStyleDef = {
  name: 'SlicerStyleOther1',
  pivot: false,
  table: false,
  elements: {
    // [measure] refs/slicer-other1.png: frame border #808080
    wholeTable: { dxf: { border: thinBorder({ rgb: 'FF808080' }) } },
    // [measure] refs/slicer-other1.png: header separator #A6A6A6 (sampled y=27)
    headerRow: {
      dxf: {
        font: { bold: true, color: { theme: 1 } },
        fill: solidFill({ rgb: 'FFFFFFFF' }),
        border: { bottom: { style: 'thin', color: { rgb: 'FFA6A6A6' } } },
      },
    },
  },
}

const slicerOther1Slicer: SlicerStyleDef = {
  name: 'SlicerStyleOther1',
  elements: {
    // [measure] refs/slicer-other1.png: solid fill #BFBFBF, border #999999
    selectedItemWithData: {
      dxf: {
        fill: solidFill({ rgb: 'FFBFBFBF' }),
        border: thinBorder({ rgb: 'FF999999' }),
        font: { color: { theme: 1 } },
      },
    },
    // [measure] refs/slicer-other1.png: white fill, border #CCCCCC
    unselectedItemWithData: {
      dxf: {
        fill: solidFill({ rgb: 'FFFFFFFF' }),
        border: thinBorder({ rgb: 'FFCCCCCC' }),
        font: { color: { theme: 1 } },
      },
    },
    // [estimate: not visible in still captures]
    selectedItemWithNoData: {
      dxf: {
        fill: solidFill({ rgb: 'FFD4D4D4' }),
        border: thinBorder({ rgb: 'FFE0E0E0' }),
        font: { color: { rgb: 'FF828282' } },
      },
    },
    // [estimate: not visible in still captures]
    unselectedItemWithNoData: {
      dxf: {
        fill: solidFill({ rgb: 'FFFFFFFF' }),
        border: thinBorder({ rgb: 'FFE0E0E0' }),
        font: { color: { rgb: 'FF828282' } },
      },
    },
    // [estimate: hover state not measurable from still captures]
    hoveredSelectedItemWithData: {
      dxf: {
        fill: solidFill({ rgb: 'FFB0B0B0' }),
        border: thinBorder({ rgb: 'FF999999' }),
        font: { color: { theme: 1 } },
      },
    },
    // [estimate: hover state not measurable from still captures]
    hoveredUnselectedItemWithData: {
      dxf: {
        fill: solidFill({ rgb: 'FFE8E8E8' }),
        border: thinBorder({ rgb: 'FFCCCCCC' }),
        font: { color: { theme: 1 } },
      },
    },
    // [estimate: hover state not measurable from still captures]
    hoveredSelectedItemWithNoData: {
      dxf: {
        fill: solidFill({ rgb: 'FFC8C8C8' }),
        border: thinBorder({ rgb: 'FFE0E0E0' }),
        font: { color: { rgb: 'FF828282' } },
      },
    },
    // [estimate: hover state not measurable from still captures]
    hoveredUnselectedItemWithNoData: {
      dxf: {
        fill: solidFill({ rgb: 'FFF5F5F5' }),
        border: thinBorder({ rgb: 'FFE0E0E0' }),
        font: { color: { rgb: 'FF828282' } },
      },
    },
  },
}

// [measure] refs/slicer-other2.png: frame border Accent 1 #4F81BD
const slicerOther2Table: TableStyleDef = {
  name: 'SlicerStyleOther2',
  pivot: false,
  table: false,
  elements: {
    // [measure] refs/slicer-other2.png: frame border Accent 1 #4F81BD
    wholeTable: { dxf: { border: thinBorder({ theme: 4 }) } },
    // [measure] refs/slicer-other2.png: header separator Accent 1 #4F81BD
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
    // [measure] refs/slicer-other2.png (flat approx of gradient AAC2E3..CBDFF6): fill #C2D8F1, border #999999
    selectedItemWithData: {
      dxf: {
        fill: solidFill({ rgb: 'FFC2D8F1' }),
        border: thinBorder({ rgb: 'FF999999' }),
        font: { color: { theme: 1 } },
      },
    },
    // [measure] refs/slicer-other2.png (flat approx of gradient EBEDF0..F8F9FA): fill #F4F5F7, border #CCCCCC
    unselectedItemWithData: {
      dxf: {
        fill: solidFill({ rgb: 'FFF4F5F7' }),
        border: thinBorder({ rgb: 'FFCCCCCC' }),
        font: { color: { theme: 1 } },
      },
    },
    // [estimate: not visible in still captures]
    selectedItemWithNoData: {
      dxf: {
        fill: solidFill({ rgb: 'FFE2EAF5' }),
        border: thinBorder({ rgb: 'FFE0E0E0' }),
        font: { color: { rgb: 'FF828282' } },
      },
    },
    // [estimate: not visible in still captures]
    unselectedItemWithNoData: {
      dxf: {
        fill: solidFill({ rgb: 'FFFFFFFF' }),
        border: thinBorder({ rgb: 'FFE0E0E0' }),
        font: { color: { rgb: 'FF828282' } },
      },
    },
    // [estimate: hover state not measurable from still captures]
    hoveredSelectedItemWithData: {
      dxf: {
        fill: solidFill({ rgb: 'FFB0CCED' }),
        border: thinBorder({ rgb: 'FF999999' }),
        font: { color: { theme: 1 } },
      },
    },
    // [estimate: hover state not measurable from still captures]
    hoveredUnselectedItemWithData: {
      dxf: {
        fill: solidFill({ rgb: 'FFE8EEF7' }),
        border: thinBorder({ rgb: 'FFCCCCCC' }),
        font: { color: { theme: 1 } },
      },
    },
    // [estimate: hover state not measurable from still captures]
    hoveredSelectedItemWithNoData: {
      dxf: {
        fill: solidFill({ rgb: 'FFD0DFF0' }),
        border: thinBorder({ rgb: 'FFE0E0E0' }),
        font: { color: { rgb: 'FF828282' } },
      },
    },
    // [estimate: hover state not measurable from still captures]
    hoveredUnselectedItemWithNoData: {
      dxf: {
        fill: solidFill({ rgb: 'FFF5F7FA' }),
        border: thinBorder({ rgb: 'FFE0E0E0' }),
        font: { color: { rgb: 'FF828282' } },
      },
    },
  },
}

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
      // [approx] vertical gradient in Excel: top B9CEE5, bottom 5489C2; flat mid-sample ~#8FB1D6
      // [measure] refs/timeline-light\${index}-months.png: dark text on white background
      selectedItemWithData: {
        dxf: {
          fill: solidFill({ theme, tint: 0.4 }),
          font: { color: { theme: 1 } },
        },
      },
      // [estimate: entire timeline period is selected in captures]
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
      // [approx] vertical gradient in Excel: top 4C84BF, bottom 214468; flat mid-sample ~#39699A
      // [measure] refs/timeline-dark\${index}-months.png: dark text on white background (not white text)
      selectedItemWithData: {
        dxf: {
          fill: solidFill({ theme, tint: -0.2 }),
          font: { color: { theme: 1 } },
        },
      },
      // [estimate: entire timeline period is selected in captures]
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

export async function generateBuiltins(cliXmlPath?: string): Promise<void> {
  const xml = await ensureXml(cliXmlPath)
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

  // Emit table-styles-digest.json for test assertions against dropped sides or mismatched element counts
  const digest: Record<string, { elementCount: number; borders: Partial<Record<TableStyleElementType, Partial<Record<string, string>>>> }> = {}
  for (const s of [...tableStyles, ...pivotStyles]) {
    digest[s.name] = {
      elementCount: s.rawElementCount,
      borders: s.borders,
    }
  }
  const digestPath = resolve(import.meta.dir, '../../tests/fixtures/xlsx-pivots/table-styles-digest.json')
  mkdirSync(dirname(digestPath), { recursive: true })
  writeFileSync(digestPath, JSON.stringify(digest, null, 2), 'utf8')
  console.log(`wrote ${digestPath} (${Object.keys(digest).length} styles)`)
}

if (import.meta.main) {
  await generateBuiltins(process.argv[2])
  console.log('Done generating built-in styles.')
}
