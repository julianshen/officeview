// @vitest-environment node
/**
 * Stage 5 Excel control interaction oracle test (Task 39, spec §9, §11).
 *
 * Verifies that the Excel-saved control interaction cases match the specification:
 * - Exact pivot cells (headers, row labels, col labels, values, totals).
 * - Slicer cache state (selected items, no-data flags, item order).
 * - Timeline cache state (filterType, date selection range, level).
 * - Markers presence in saved XML.
 * - Sensitivity: rejects missing, wrong, or unchanged actions.
 * - Idempotency and round-trip symmetry (clear restores seed, level change preserves range).
 */
import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { CONTROL_INTERACTION_CASES, type ControlInteractionCase } from '../scripts/xlsx-refs/control-interaction-cases'
import { controlsSeedVerified, isCaseOutputVerified } from '../scripts/xlsx-refs/build-control-interactions'
import { pivotFieldXmlScope } from '../scripts/xlsx-refs/control-interaction-xml'

const ROOT = fileURLToPath(new URL('..', import.meta.url))
const CONTROLS_DIR = resolve(ROOT, 'tests/fixtures/xlsx-pivots/controls')
const MANIFEST_PATH = resolve(CONTROLS_DIR, 'manifest.json')
const SEED_PATH = resolve(CONTROLS_DIR, 'controls-seed.xlsx')

async function extractCells(file: string, sheetPath = 'xl/worksheets/sheet1.xml'): Promise<Record<string, string | number | null>> {
  const zip = await JSZip.loadAsync(readFileSync(file))
  const sheetXml = await zip.file(sheetPath)!.async('string')
  let strings: string[] = []
  const ssFile = zip.file('xl/sharedStrings.xml')
  if (ssFile) {
    const ssXml = await ssFile.async('string')
    const matches = ssXml.matchAll(/<si>(?:(?!<\/si>).)*?<t[^>]*>(.*?)<\/t>/gs)
    strings = Array.from(matches, m => m[1] ?? '')
  }
  const cells: Record<string, string | number | null> = {}
  const cellMatches = sheetXml.matchAll(/<c r="([A-Z0-9]+)"(?:\s+s="[^"]*")?(?:\s+t="([^"]*)")?>.*?<v>(.*?)<\/v><\/c>/gs)
  for (const m of cellMatches) {
    const r = m[1]!
    const t = m[2]
    const rawVal = m[3]!
    if (t === 's') {
      const idx = Number(rawVal)
      cells[r] = strings[idx] ?? rawVal
    } else if (t === 'inlineStr') {
      const strMatch = /<t[^>]*>(.*?)<\/t>/.exec(m[0])
      cells[r] = strMatch ? strMatch[1]! : rawVal
    } else {
      const num = Number(rawVal)
      cells[r] = Number.isFinite(num) ? num : rawVal
    }
  }
  return cells
}

function cellPosition(cell: string): { row: number; col: number } {
  const match = /^([A-Z]+)(\d+)$/.exec(cell)
  if (!match) throw new Error(`Invalid A1 cell ${cell}`)
  let col = 0
  for (const char of match[1]!) col = col * 26 + char.charCodeAt(0) - 64
  return { row: Number(match[2]), col }
}

async function pivotCellMap(zip: JSZip, pivotName: string, cells: Record<string, string | number | null>): Promise<Record<string, string | number | null>> {
  const pivotParts = Object.keys(zip.files).filter(path => /^xl\/pivotTables\/pivotTable[^/]+\.xml$/.test(path))
  const matches: Array<{ part: string; ref: string }> = []
  for (const part of pivotParts) {
    const xml = await zip.file(part)!.async('string')
    const name = /<pivotTableDefinition\b[^>]*\bname="([^"]+)"/.exec(xml)?.[1]
    const ref = /<location\b[^>]*\bref="([A-Z]+\d+:[A-Z]+\d+)"/.exec(xml)?.[1]
    if (name === pivotName && ref) matches.push({ part, ref })
  }
  if (matches.length !== 1) throw new Error(`Expected one pivot ${pivotName}, found ${matches.length}`)
  const [topLeft, bottomRight] = matches[0]!.ref.split(':').map(cellPosition)
  return Object.fromEntries(Object.entries(cells).filter(([cell]) => {
    const { row, col } = cellPosition(cell)
    return row >= topLeft!.row && row <= bottomRight!.row && col >= topLeft!.col && col <= bottomRight!.col
  }))
}

function slicerItems(xml: string): Array<{ x: number; selected: boolean; noData: boolean }> {
  const block = /<items\b[^>]*>([\s\S]*?)<\/items>/.exec(xml)?.[1]
  if (block === undefined) throw new Error('slicer cache has no items block')
  return [...block.matchAll(/<i\b([^>]*)\/>/g)].map(match => {
    const attrs = match[1] ?? ''
    const x = /\bx="(\d+)"/.exec(attrs)?.[1]
    if (x === undefined) throw new Error('slicer item is missing its x id')
    return { x: Number(x), selected: /\bs="1"/.test(attrs), noData: /\bnd="1"/.test(attrs) }
  })
}

describe('Stage 5 Excel Control Oracle (Task 39)', () => {
  it('controls-seed.xlsx exists and has the expected seed structure', async () => {
    expect(existsSync(SEED_PATH), 'controls-seed.xlsx must exist').toBe(true)
    const seedCells = await extractCells(SEED_PATH)
    // Pivot header and grand total
    expect(seedCells.A3).toBe('加總 - Sales')
    expect(seedCells.A4).toBe('列標籤')
    expect(seedCells.B4).toBe('Apples')
    expect(seedCells.C4).toBe('Bananas')
    expect(seedCells.D4).toBe('Cherries')
    expect(seedCells.E4).toBe('總計')
    expect(seedCells.A8).toBe('總計')
    expect(seedCells.E8).toBe(150)
  })

  it('manifest.json matches CONTROL_INTERACTION_CASES', () => {
    expect(existsSync(MANIFEST_PATH), 'manifest.json must exist').toBe(true)
    const manifest = JSON.parse(readFileSync(MANIFEST_PATH, 'utf8')) as ControlInteractionCase[]
    expect(manifest.length).toBe(CONTROL_INTERACTION_CASES.length)
    for (let i = 0; i < manifest.length; i++) {
      expect(manifest[i]!.case).toBe(CONTROL_INTERACTION_CASES[i]!.case)
      expect(manifest[i]!.status).toBe(CONTROL_INTERACTION_CASES[i]!.status)
      expect(manifest[i]!.action).toEqual(CONTROL_INTERACTION_CASES[i]!.action)
      expect(manifest[i]!.inputMarkers).toEqual(CONTROL_INTERACTION_CASES[i]!.inputMarkers)
      expect(manifest[i]!.scopedMarkers).toEqual(CONTROL_INTERACTION_CASES[i]!.scopedMarkers)
      expect(manifest[i]!).not.toHaveProperty('markers')
    }
  })

  it('resumes only from seed and case outputs that still match their oracles', async () => {
    expect(await controlsSeedVerified(SEED_PATH)).toBe(true)
    for (const c of CONTROL_INTERACTION_CASES.filter(testCase => testCase.status === 'built')) {
      expect(await isCaseOutputVerified(c, resolve(ROOT, c.file)), c.case).toBe(true)
    }
    const pending = CONTROL_INTERACTION_CASES.filter(testCase => testCase.status === 'unverified')
    for (const c of pending) {
      expect(await isCaseOutputVerified(c, resolve(ROOT, c.file)), c.case).toBe(false)
    }
  })

  describe('Built cases verify against Excel-saved workbooks', () => {
    const runnableCases = CONTROL_INTERACTION_CASES.filter(c => c.status === 'built' || (c.route === 'xml' && existsSync(resolve(ROOT, c.file))))
    const pendingCases = CONTROL_INTERACTION_CASES.filter(c => !runnableCases.includes(c))

    for (const c of pendingCases) {
      it.skip(`SKIPPED: ${c.case} (${c.status} route: ${c.route}; see recipe notes or run the Excel builder)`, () => {})
    }

    it.each(runnableCases.map(c => [c.case, c] as const))('%s', async (_name, c) => {
      const filePath = resolve(ROOT, c.file)
      expect(existsSync(filePath), `file ${c.file} must exist`).toBe(true)

      const zip = await JSZip.loadAsync(readFileSync(filePath))
      const actualCells = await extractCells(filePath)

      // 1. Pivot cell values match expected
      for (const pivot of c.expected.pivots ?? []) {
        expect(await pivotCellMap(zip, pivot.pivotName, actualCells), `${c.case} ${pivot.pivotName} complete location map`).toEqual(pivot.cells)
        for (const [coord, val] of Object.entries(pivot.cells)) {
          expect(actualCells[coord], `${c.case} ${pivot.pivotName} cell ${coord}`).toBe(val)
        }
        for (const coord of pivot.absentCells ?? []) {
          expect(actualCells[coord], `${c.case} ${pivot.pivotName} cell ${coord} should be absent`).toBeUndefined()
        }
      }

      // 2. Slicer cache verification if present
      if (c.expected.slicerCache) {
        const cachePart = c.expected.slicerCache.part ?? 'xl/slicerCaches/slicerCache1.xml'
        const scXml = await zip.file(cachePart)?.async('string')
        expect(scXml, `${c.case} must contain ${cachePart}`).toBeDefined()
        expect(slicerItems(scXml!)).toEqual(c.expected.slicerCache.items.map(item => ({ x: item.x, selected: item.selected, noData: item.noData ?? false })))
      }

      // Secondary slicer cache if present (e.g. crossfilter)
      if (c.expected.secondarySlicerCache) {
        const scXml = await zip.file(c.expected.secondarySlicerCache.part ?? 'xl/slicerCaches/slicerCache1.xml')?.async('string')
        expect(scXml, `${c.case} must contain secondary slicerCache1.xml`).toBeDefined()
        expect(slicerItems(scXml!)).toEqual(c.expected.secondarySlicerCache.items.map(item => ({ x: item.x, selected: item.selected, noData: item.noData ?? false })))
      }

      // 3. Timeline cache verification if present
      if (c.expected.timelineCache) {
        const tcXml = await zip.file('xl/timelineCaches/timelineCache1.xml')?.async('string')
        expect(tcXml, `${c.case} must contain timelineCache1.xml`).toBeDefined()
        if (c.expected.timelineCache.filterType) {
          expect(tcXml).toContain(`filterType="${c.expected.timelineCache.filterType}"`)
        }
        if (c.expected.timelineCache.selection) {
          expect(tcXml).toContain(`startDate="${c.expected.timelineCache.selection.startDate}"`)
          expect(tcXml).toContain(`endDate="${c.expected.timelineCache.selection.endDate}"`)
        }
      }

      // 4. Table rows verification if present (for s5-table-slicer)
      if (c.expected.tableRows) {
        const s2Xml = await zip.file('xl/worksheets/sheet2.xml')?.async('string')
        expect(s2Xml).toBeDefined()
        for (let idx = 1; idx <= 5; idx++) {
          const rowNum = idx + 1
          const isVisible = c.expected.tableRows.visibleRowIndices.includes(idx)
          const rowRegex = new RegExp(`<row r="${rowNum}"[^>]*hidden="1"`)
          if (isVisible) {
            expect(s2Xml).not.toMatch(rowRegex)
          } else {
            expect(s2Xml).toMatch(rowRegex)
          }
        }
      }

      // 5. Scoped markers verification
      if (c.scopedMarkers?.length) {
        for (const sm of c.scopedMarkers) {
          const partXml = await zip.file(sm.part)?.async('string')
          expect(partXml, `part ${sm.part} must exist in ${c.case}`).toBeDefined()
          const markerXml = sm.within ? pivotFieldXmlScope(partXml!, sm.within.value) : partXml!
          if (sm.match) expect(markerXml.split(sm.match).length - 1, `${c.case} part ${sm.part}${sm.within ? ` within ${sm.within.attribute}=${sm.within.value}` : ''} must contain one "${sm.match}"`).toBe(1)
          if (sm.absent) {
            expect(markerXml, `${c.case} part ${sm.part}${sm.within ? ` within ${sm.within.attribute}=${sm.within.value}` : ''} must NOT contain "${sm.absent}"`).not.toContain(sm.absent)
          }
        }
      }
    })
  })

  describe('Semantics, symmetry and sensitivity', () => {
    it.skipIf(!existsSync(resolve(CONTROLS_DIR, 's5-timeline-click-month.xlsx')))('s5-timeline-click-month is March 2024 inclusive of March 31 and excludes April 1', async () => {
      const cells = await extractCells(resolve(CONTROLS_DIR, 's5-timeline-click-month.xlsx'))
      expect(cells.B4).toBe('Cherries')
      expect(cells.B5).toBe(40)
      expect(cells.C5).toBe(40)
      expect(cells.A5).toBe('West')
      expect(cells.A6).toBe('總計')
      expect(cells.D4).toBeUndefined()
    })

    it.skipIf(!existsSync(resolve(CONTROLS_DIR, 's5-timeline-level-preserves-range.xlsx')))('s5-timeline-level-preserves-range includes Jan through Mar, excludes Dec 31 and Apr 1, and totals 90', async () => {
      const cells = await extractCells(resolve(CONTROLS_DIR, 's5-timeline-level-preserves-range.xlsx'))
      expect(cells.C5).toBe(20)
      expect(cells.E5).toBe(20)
      expect(cells.E6).toBe(70)
      expect(cells.B7).toBe(30)
      expect(cells.C7).toBe(20)
      expect(cells.D7).toBe(40)
      expect(cells.E7).toBe(90)
      expect(cells.E8).toBeUndefined()
    })

    it.skipIf(!existsSync(resolve(CONTROLS_DIR, 's5-slicer-clear.xlsx')))('s5-slicer-clear restores the full pivot table values from controls-seed.xlsx', async () => {
      const seedCells = await extractCells(SEED_PATH)
      const clearCells = await extractCells(resolve(CONTROLS_DIR, 's5-slicer-clear.xlsx'))
      expect(clearCells).toEqual(seedCells)
    })

    it.skipIf(!existsSync(resolve(CONTROLS_DIR, 's5-timeline-clear.xlsx')))('s5-timeline-clear restores the full pivot table values from controls-seed.xlsx', async () => {
      const seedCells = await extractCells(SEED_PATH)
      const clearCells = await extractCells(resolve(CONTROLS_DIR, 's5-timeline-clear.xlsx'))
      expect(clearCells).toEqual(seedCells)
    })

    it.skipIf(!existsSync(resolve(CONTROLS_DIR, 's5-timeline-level-preserves-range.xlsx')))('s5-timeline-level-preserves-range retains identical cells to drag-forward-reverse', async () => {
      const dragCells = await extractCells(resolve(CONTROLS_DIR, 's5-timeline-drag-forward-reverse.xlsx'))
      const levelCells = await extractCells(resolve(CONTROLS_DIR, 's5-timeline-level-preserves-range.xlsx'))
      expect(levelCells).toEqual(dragCells)
    })

    it.skipIf(!existsSync(resolve(CONTROLS_DIR, 's5-linked-slicer-and-timeline.xlsx')))('s5-linked-slicer-and-timeline produces empty intersection without crash or data rows', async () => {
      const cells = await extractCells(resolve(CONTROLS_DIR, 's5-linked-slicer-and-timeline.xlsx'))
      expect(cells.A3).toBe('加總 - Sales')
      expect(cells.B4).toBe('總計')
      expect(cells.A5).toBe('總計')
      // No region rows (East, South, West)
      expect(cells.A6).toBeUndefined()
      expect(cells.B5).toBeUndefined()
    })

    it('s5-slicer-manual-filter-vs-slicer records Excel restoring the selected Cherries item', async () => {
      const cells = await extractCells(resolve(CONTROLS_DIR, 's5-slicer-manual-filter-vs-slicer.xlsx'))
      expect(cells.D4).toBe('Cherries')
      expect(cells.E8).toBe(150)
    })

    it('detects and rejects missing or mismatched cell changes (sensitivity gate)', async () => {
      const seedCells = await extractCells(SEED_PATH)
      const clickOneCells = await extractCells(resolve(CONTROLS_DIR, 's5-slicer-click-one.xlsx'))
      // Must not match seed (action actually changed state)
      expect(clickOneCells).not.toEqual(seedCells)
      // Grand total filtered from 150 to 90
      expect(seedCells.E8).toBe(150)
      expect(clickOneCells.C8).toBe(90)
    })
  })
})
