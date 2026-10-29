/**
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
    // [approx] vertical gradient in Excel (AAC2E3..CBDFF6): flat mid-fill #C2D8F1, border #999999
    selectedItemWithData: {
      dxf: {
        fill: solidFill({ rgb: 'FFC2D8F1' }),
        border: thinBorder({ rgb: 'FF999999' }),
        font: { color: { theme: 1 } },
      },
    },
    // [approx] vertical gradient in Excel (EBEDF0..F8F9FA): flat mid-fill #F4F5F7, border #CCCCCC
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
  const name = `TimeSlicerStyleLight${index}`
  const tableDef: TableStyleDef = {
    name,
    pivot: false,
    table: false,
    elements: {
      // [measure] refs/timeline-light${index}-months.png: frame border in style accent (1 px line anti-aliased in capture)
      wholeTable: { dxf: { border: thinBorder({ theme }) } },
      // [measure] refs/timeline-light${index}-months.png: header separator in style accent (1 px line anti-aliased in capture)
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
      // [measure] refs/timeline-light${index}-months.png: period labels #909090 on white (year labels #989898)
      selectedItemWithData: {
        dxf: {
          fill: solidFill({ theme, tint: 0.4 }),
          font: { color: { rgb: 'FF909090' } },
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
  const name = `TimeSlicerStyleDark${index}`
  const tableDef: TableStyleDef = {
    name,
    pivot: false,
    table: false,
    elements: {
      // [measure] refs/timeline-dark${index}-months.png: frame border in style accent (1 px line anti-aliased in capture)
      wholeTable: { dxf: { border: thinBorder({ theme }) } },
      // [measure] refs/timeline-dark${index}-months.png: header separator in style accent (1 px line anti-aliased in capture)
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
      // [measure] refs/timeline-dark${index}-months.png: period labels #909090 on white (year labels #989898)
      selectedItemWithData: {
        dxf: {
          fill: solidFill({ theme, tint: -0.2 }),
          font: { color: { rgb: 'FF909090' } },
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
