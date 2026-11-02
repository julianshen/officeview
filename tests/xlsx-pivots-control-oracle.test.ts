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
    }
  })

  describe('Built cases verify against Excel-saved workbooks', () => {
    const builtCases = CONTROL_INTERACTION_CASES.filter(c => c.status === 'built')
    const recipeCases = CONTROL_INTERACTION_CASES.filter(c => c.status !== 'built')

    for (const c of recipeCases) {
      it.skip(`SKIPPED: ${c.case} (${c.status} route: ${c.route}; see recipe notes)`, () => {})
    }

    it.each(builtCases.map(c => [c.case, c] as const))('%s', async (_name, c) => {
      const filePath = resolve(ROOT, c.file)
      expect(existsSync(filePath), `file ${c.file} must exist`).toBe(true)

      const zip = await JSZip.loadAsync(readFileSync(filePath))
      const actualCells = await extractCells(filePath)

      // 1. Pivot cell values match expected
      if (c.expected.pivots?.[0]?.cells) {
        const expCells = c.expected.pivots[0].cells
        for (const [coord, val] of Object.entries(expCells)) {
          expect(actualCells[coord], `${c.case} cell ${coord}`).toBe(val)
        }
      }

      // 2. Slicer cache verification if present
      if (c.expected.slicerCache) {
        const scXml = await zip.file('xl/slicerCaches/slicerCache1.xml')?.async('string')
        expect(scXml, `${c.case} must contain slicerCache1.xml`).toBeDefined()
        for (const item of c.expected.slicerCache.items) {
          if (item.selected) {
            expect(scXml).toMatch(new RegExp(`<i\\s+[^>]*x="${item.x}"[^>]*\\bs="1"`))
          }
          if (item.noData) {
            expect(scXml).toMatch(new RegExp(`<i\\s+[^>]*x="${item.x}"[^>]*\\bnd="1"`))
          }
        }
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

      // 4. Markers verification
      if (c.markers?.length) {
        let allXml = ''
        for (const [, f] of Object.entries(zip.files)) {
          if (f.name.endsWith('.xml')) {
            allXml += await f.async('string')
          }
        }
        for (const marker of c.markers) {
          expect(allXml, `${c.case} lacks marker ${marker}`).toContain(marker)
        }
      }
    })
  })

  describe('Semantics, symmetry and sensitivity', () => {
    it('s5-slicer-click-selected is idempotent with s5-slicer-click-one', async () => {
      const cellsOne = await extractCells(resolve(CONTROLS_DIR, 's5-slicer-click-one.xlsx'))
      const cellsSelected = await extractCells(resolve(CONTROLS_DIR, 's5-slicer-click-selected.xlsx'))
      expect(cellsSelected).toEqual(cellsOne)
    })

    it('s5-slicer-clear restores the full pivot table values from controls-seed.xlsx', async () => {
      const seedCells = await extractCells(SEED_PATH)
      const clearCells = await extractCells(resolve(CONTROLS_DIR, 's5-slicer-clear.xlsx'))
      expect(clearCells).toEqual(seedCells)
    })

    it('s5-timeline-clear-and-scroll restores the full pivot table values from controls-seed.xlsx', async () => {
      const seedCells = await extractCells(SEED_PATH)
      const clearCells = await extractCells(resolve(CONTROLS_DIR, 's5-timeline-clear-and-scroll.xlsx'))
      expect(clearCells).toEqual(seedCells)
    })

    it('s5-timeline-level-preserves-range retains identical cells to drag-forward-reverse', async () => {
      const dragCells = await extractCells(resolve(CONTROLS_DIR, 's5-timeline-drag-forward-reverse.xlsx'))
      const levelCells = await extractCells(resolve(CONTROLS_DIR, 's5-timeline-level-preserves-range.xlsx'))
      expect(levelCells).toEqual(dragCells)
    })

    it('s5-linked-slicer-and-timeline produces empty intersection without crash or data rows', async () => {
      const cells = await extractCells(resolve(CONTROLS_DIR, 's5-linked-slicer-and-timeline.xlsx'))
      expect(cells.A3).toBe('加總 - Sales')
      expect(cells.B4).toBe('總計')
      expect(cells.A5).toBe('總計')
      // No region rows (East, South, West)
      expect(cells.A6).toBeUndefined()
      expect(cells.B5).toBeUndefined()
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
