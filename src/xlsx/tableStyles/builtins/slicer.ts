/**
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
  const name = `SlicerStyleLight${index}`
  const tableDef: TableStyleDef = {
    name,
    pivot: false,
    table: false,
    elements: {
      // [measure] refs/slicer-light${index}.png: frame border in style accent
      wholeTable: { dxf: { border: thinBorder({ theme }) } },
      // [measure] refs/slicer-light${index}.png: header on white with accent separator
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
      // [measure] refs/slicer-light${index}.png: selected item fill theme tint 0.6, border #999999
      selectedItemWithData: {
        dxf: {
          fill: solidFill({ theme, tint: 0.6 }),
          border: thinBorder({ rgb: 'FF999999' }),
          font: { color: { theme: 1 } },
        },
      },
      // [measure] refs/slicer-light${index}.png: unselected item fill white, border #CCCCCC
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
  const name = `SlicerStyleDark${index}`
  const tableDef: TableStyleDef = {
    name,
    pivot: false,
    table: false,
    elements: {
      // [measure] refs/slicer-dark${index}.png: frame border in style accent
      wholeTable: { dxf: { border: thinBorder({ theme }) } },
      // [measure] refs/slicer-dark${index}.png: header on white with accent separator
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
      // [measure] refs/slicer-dark${index}.png: selected item solid accent fill, matching border, bold white text
      selectedItemWithData: {
        dxf: {
          fill: solidFill({ theme }),
          border: thinBorder({ theme }),
          font: { bold: true, color: { theme: 0 } },
        },
      },
      // [measure] refs/slicer-dark${index}.png: unselected item silver #C0C0C0 fill and border, dark text
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
  const name = `TimeSlicerStyleLight${index}`
  const tableDef: TableStyleDef = {
    name,
    pivot: false,
    table: false,
    elements: {
      // [measure] refs/timeline-light${index}-months.png: frame border in style accent
      wholeTable: { dxf: { border: thinBorder({ theme }) } },
      // [measure] refs/timeline-light${index}-months.png: header separator in style accent
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
      // [measure] refs/timeline-light${index}-months.png: selected period bar light accent fill
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
  const name = `TimeSlicerStyleDark${index}`
  const tableDef: TableStyleDef = {
    name,
    pivot: false,
    table: false,
    elements: {
      // [measure] refs/timeline-dark${index}-months.png: frame border in style accent
      wholeTable: { dxf: { border: thinBorder({ theme }) } },
      // [measure] refs/timeline-dark${index}-months.png: header separator in style accent
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
      // [measure] refs/timeline-dark${index}-months.png: selected period bar dark accent shade fill
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
