/**
 * S10.1 browser fixtures and drivers for the real interactive viewer page
 * (src/debug/viewer-e2e.html). Packages are true XLSX files: the committed S0.2
 * small packages, or ones written here by the S0.2 builders at an explicit
 * (and labelled) scale. The page fetches /__fixtures/<name>, which each test
 * routes to the file on disk, so nothing outside src/debug is served by vite.
 */
import { expect, type Page } from '@playwright/test'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { writeFixturePackage, writeSheetsPackage, type SheetRowSpec } from '../../scripts/generate-viewer-fixtures'
import { VIEWER_E2E_FIXTURE_DIR as GENERATED } from '../../scripts/harness/viewerE2EFixtureCache'
import type { CellRef, FixtureCell } from '../helpers/xlsx-viewer-fixtures'
import { xlsxBytes } from '../helpers/xlsx-package'
import { structureBytes, twoCell } from '../helpers/xlsx-structure-fixtures'

const ROOT = fileURLToPath(new URL('../..', import.meta.url))
const COMMITTED = join(ROOT, 'tests', 'fixtures', 'xlsx-viewer')

/** Generated packages: name → (scale label, writer). Object caps for profile D
 * are 100,000 rows / 100,000 populated cells, so the "paged" packages cross one
 * cap and open worker-backed when the page supplies a viewer worker. */
export const GENERATED_FIXTURES = {
  'f1-paged.xlsx': 'F1 sampled corners, nnz 120,000 (crosses the 100,000-cell object cap)',
  'f2-1k.xlsx': 'F2 tall table at 1,000 rows (reduced scale, object path)',
  'f2-paged.xlsx': 'F2 tall table at 100,010 rows (reduced scale; crosses the 100,000-row object cap)',
  'f2-storage.xlsx': 'F2 tall table at 160,000 rows (reduced scale; its expected page bytes exceed the 32 MiB in-memory page store, so it needs OPFS/IndexedDB)',
  'f3-paged.xlsx': 'F3 dense wide table at 8 rows x 16,384 columns (reduced rows; 131,072 cells)',
  'f2-sort8.xlsx': 'F2 variant: 8 columns with tied prefix groups for 8-key multi-level sort acceptance',
  'f4-12k.xlsx': 'F4 string domain at 12,000 rows (reduced scale; > 10,000 unique values)',
  'f5-structure.xlsx': 'F5-shaped structure: giant narrow-column merge D1:D1048576, A1:B2 merge, fixed array C30:C32',
  'f5-spill.xlsx': 'F5 structure variant: dynamic array spill D20:D23 (SEQUENCE)',
  'f5-drawing.xlsx': 'F5 structure variant: twoCell drawing anchor at rows 10..13',
  'f7-two-sheets.xlsx': 'F7-shaped two sheets A/B with identical first revisions, hidden rows and a tiny row height',
  'f7-three-sheets.xlsx': 'F7 variant: three sheets A/B/C for multi-sheet switching memory plateau',
} as const
export type FixtureName = keyof typeof GENERATED_FIXTURES | 'f1-small.xlsx' | 'f3sparse-small.xlsx' | 'f4-small.xlsx' | 'f5-small.xlsx' | 'f8-small.xlsx'
  /** S13: written by native Microsoft Excel (see tests/xlsx-viewer-native-excel.test.ts). */
  | 'native-corner.xlsx' | 'native-1904.xlsx' | 'large/native-large.xlsx'

const num = (value: number, styleIndex = 0): FixtureCell => ({ value, hasCachedValue: true, valueIsError: false, styleIndex })
const text = (value: string, styleIndex = 0): FixtureCell => ({ value, hasCachedValue: true, valueIsError: false, styleIndex })

/** Sort table: header (id, score, label), `rows` body rows; score descends with id. */
function sortTable(rows: number, offset: number): SheetRowSpec[] {
  const out: SheetRowSpec[] = [{ r: 0, cells: [{ row: 0, col: 0, cell: text('id', 1) }, { row: 0, col: 1, cell: text('score', 1) }, { row: 0, col: 2, cell: text('label', 1) }] }]
  for (let r = 1; r <= rows; r++) {
    const cells: CellRef[] = [{ row: r, col: 0, cell: num(r) }, { row: r, col: 1, cell: num(offset + rows - r) }, { row: r, col: 2, cell: text(`s${offset}-${r}`) }]
    out.push({ r, cells, ...(r === 7 ? { hidden: true } : {}), ...(r === 9 ? { ht: 0.75 } : {}) })
  }
  return out
}

function f5Package(): Uint8Array {
  const rows: string[] = []
  rows.push('<row r="1"><c r="A1" t="inlineStr"><is><t>key</t></is></c><c r="B1" t="inlineStr"><is><t>val</t></is></c><c r="D1"><v>1</v></c></row>')
  for (let r = 2; r <= 40; r++) {
    const spill = r === 30 ? `<c r="C30"><f t="array" ref="C30:C32">ROW(A1:A3)</f><v>1</v></c>` : r === 31 ? '<c r="C31"><v>2</v></c>' : r === 32 ? '<c r="C32"><v>3</v></c>' : ''
    rows.push(`<row r="${r}"><c r="A${r}"><v>${(r * 37) % 41}</v></c><c r="B${r}"><v>${r}</v></c>${spill}</row>`)
  }
  return xlsxBytes({ sheetData: rows.join(''), afterSheetData: '<mergeCells count="1"><mergeCell ref="D1:D1048576"/></mergeCells>' })
}

function f5SpillPackage(): Uint8Array {
  const rows: string[] = []
  rows.push('<row r="1"><c r="A1" t="inlineStr"><is><t>key</t></is></c><c r="B1" t="inlineStr"><is><t>val</t></is></c></row>')
  for (let r = 2; r <= 30; r++) {
    const spill = r === 20
      ? '<c r="D20" cm="1"><f t="array" ref="D20:D23">_xlfn.SEQUENCE(4)</f><v>1</v></c>'
      : r >= 21 && r <= 23
        ? `<c r="D${r}"><v>${r - 19}</v></c>`
        : ''
    rows.push(`<row r="${r}"><c r="A${r}"><v>${31 - r}</v></c><c r="B${r}"><v>${r}</v></c>${spill}</row>`)
  }
  return structureBytes({
    dynamic: true,
    sheetData: rows.join(''),
  })
}

function f5DrawingPackage(): Uint8Array {
  const rows: string[] = []
  rows.push('<row r="1"><c r="A1" t="inlineStr"><is><t>key</t></is></c><c r="B1" t="inlineStr"><is><t>val</t></is></c></row>')
  for (let r = 2; r <= 30; r++) {
    const valB = r < 10 ? 'top' : r <= 20 ? 'mid' : 'bot'
    rows.push(`<row r="${r}"><c r="A${r}"><v>${31 - r}</v></c><c r="B${r}" t="inlineStr"><is><t>${valB}</t></is></c></row>`)
  }
  return structureBytes({
    sheetData: rows.join(''),
    drawing: twoCell([3, 9], [5, 12], 3),
  })
}

function f2Sort8Package(): Uint8Array {
  const rows: string[] = []
  const cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
  rows.push('<row r="1">' + cols.map(c => `<c r="${c}1" t="inlineStr"><is><t>col${c}</t></is></c>`).join('') + '<c r="I1" t="inlineStr"><is><t>label</t></is></c></row>')
  const defs: Array<{ r: number; vals: number[]; label: string }> = [
    { r: 2, vals: [1, 1, 5, 5, 5, 5, 5, 5], label: 'B-low' },
    { r: 3, vals: [1, 2, 5, 5, 5, 5, 5, 5], label: 'B-high' },
    { r: 4, vals: [2, 5, 2, 5, 5, 5, 5, 5], label: 'C-high' },
    { r: 5, vals: [2, 5, 1, 5, 5, 5, 5, 5], label: 'C-low' },
    { r: 6, vals: [3, 5, 5, 1, 5, 5, 5, 5], label: 'D-low' },
    { r: 7, vals: [3, 5, 5, 2, 5, 5, 5, 5], label: 'D-high' },
    { r: 8, vals: [4, 5, 5, 5, 2, 5, 5, 5], label: 'E-high' },
    { r: 9, vals: [4, 5, 5, 5, 1, 5, 5, 5], label: 'E-low' },
    { r: 10, vals: [5, 5, 5, 5, 5, 1, 5, 5], label: 'F-low' },
    { r: 11, vals: [5, 5, 5, 5, 5, 2, 5, 5], label: 'F-high' },
    { r: 12, vals: [6, 5, 5, 5, 5, 5, 2, 5], label: 'G-high' },
    { r: 13, vals: [6, 5, 5, 5, 5, 5, 1, 5], label: 'G-low' },
    { r: 14, vals: [7, 5, 5, 5, 5, 5, 5, 1], label: 'H-low' },
    { r: 15, vals: [7, 5, 5, 5, 5, 5, 5, 2], label: 'H-high' },
    { r: 16, vals: [9, 5, 5, 5, 5, 5, 5, 5], label: 'A-high' },
    { r: 17, vals: [8, 5, 5, 5, 5, 5, 5, 5], label: 'A-low' },
    { r: 18, vals: [10, 1, 1, 5, 5, 5, 5, 5], label: 'BC-Cwins' },
    { r: 19, vals: [10, 2, 2, 5, 5, 5, 5, 5], label: 'BC-Bwins' },
  ]
  for (const d of defs) {
    const cells = d.vals.map((v, i) => `<c r="${cols[i]}${d.r}"><v>${v}</v></c>`).join('')
    rows.push(`<row r="${d.r}">${cells}<c r="I${d.r}" t="inlineStr"><is><t>${d.label}</t></is></c></row>`)
  }
  return xlsxBytes({ sheetData: rows.join('') })
}

let generated = false
/** Writes every generated package once per machine-run (deterministic bytes). */
export async function ensureFixtures(): Promise<void> {
  if (generated) return
  mkdirSync(GENERATED, { recursive: true })
  const want = async (name: keyof typeof GENERATED_FIXTURES, write: (path: string) => Promise<unknown> | void) => {
    const path = join(GENERATED, name)
    if (!existsSync(path)) await write(path)
  }
  await want('f1-paged.xlsx', path => writeFixturePackage(path, 'F1', { nnz: 120_000 }))
  await want('f2-1k.xlsx', path => writeFixturePackage(path, 'F2', { rowCount: 1000 }))
  await want('f2-sort8.xlsx', path => writeFileSync(path, f2Sort8Package()))
  await want('f2-paged.xlsx', path => writeFixturePackage(path, 'F2', { rowCount: 100_010 }))
  await want('f2-storage.xlsx', path => writeFixturePackage(path, 'F2', { rowCount: 160_000 }))
  await want('f3-paged.xlsx', path => writeFixturePackage(path, 'F3', { rowCount: 8 }))
  await want('f4-12k.xlsx', path => writeFixturePackage(path, 'F4', { rowCount: 12_000 }))
  await want('f5-structure.xlsx', path => writeFileSync(path, f5Package()))
  await want('f5-spill.xlsx', path => writeFileSync(path, f5SpillPackage()))
  await want('f5-drawing.xlsx', path => writeFileSync(path, f5DrawingPackage()))
  await want('f7-two-sheets.xlsx', path => writeSheetsPackage(path, [{ name: 'A', rows: sortTable(40, 0) }, { name: 'B', rows: sortTable(40, 1000) }]))
  await want('f7-three-sheets.xlsx', path => writeSheetsPackage(path, [{ name: 'A', rows: sortTable(40, 0) }, { name: 'B', rows: sortTable(40, 1000) }, { name: 'C', rows: sortTable(40, 2000) }]))
  generated = true
}

export function fixturePath(name: FixtureName): string {
  return name in GENERATED_FIXTURES ? join(GENERATED, name) : join(COMMITTED, name)
}

export interface OpenOptions {
  worker?: boolean
  copy?: boolean
  /** Capabilities the viewer worker denies itself (src/debug/workerCaps.ts). */
  workerCaps?: readonly string[]
  /** With `worker`: debug-only forced paged admission (src/debug/forcePaged.ts), so a small file opens worker-backed. */
  paged?: boolean
  /** false: return once the page has navigated (capability-failure scenarios). */
  ready?: boolean
}
/** Opens a fixture in the real viewer page and waits for the record status. */
export async function openViewer(page: Page, name: FixtureName, options: OpenOptions = {}): Promise<void> {
  await page.route(url => url.pathname.startsWith('/__fixtures/'), route => {
    const requested = decodeURIComponent(new URL(route.request().url()).pathname.slice('/__fixtures/'.length)) as FixtureName
    return route.fulfill({ body: readFileSync(fixturePath(requested)), contentType: 'application/octet-stream' })
  })
  const query = new URLSearchParams({ src: `/__fixtures/${name}` })
  if (options.worker) query.set('worker', '1')
  if (options.copy === false) query.set('copy', '0')
  if (options.paged) query.set('paged', '1')
  if (options.workerCaps?.length) query.set('wcaps', options.workerCaps.join(','))
  await page.goto(`/src/debug/viewer-e2e.html?${query}`)
  if (options.ready === false) return
  await expect(page.getByRole('grid')).toBeVisible({ timeout: 60_000 })
  await expect(status(page)).toContainText('Visible body', { timeout: 60_000 })
}

/** The grid's record/status line (GridSemantics status region). */
export const status = (page: Page) => page.getByRole('status').filter({ hasText: 'Visible body' })

export async function goTo(page: Page, address: string, space: 'display' | 'source' = 'display'): Promise<void> {
  await page.getByRole('combobox', { name: 'Address space' }).selectOption(space)
  await page.getByRole('textbox', { name: 'Go To cell' }).fill(address)
  await page.getByRole('button', { name: 'Go To' }).click()
}

/** Accessible name of the grid's current active descendant (or null). */
export async function activeCell(page: Page): Promise<{ id: string; label: string } | null> {
  return page.evaluate(() => {
    const grid = document.querySelector('[role=grid]')
    const id = grid?.getAttribute('aria-activedescendant')
    const cell = id ? document.getElementById(id) : null
    return cell ? { id: cell.id, label: cell.textContent ?? '' } : null
  })
}

export const cellIdOf = (id: string) => Number(/-cell-(\d+)$/.exec(id)?.[1] ?? NaN)

export async function counts(page: Page): Promise<Record<string, number>> {
  return page.evaluate(() => window.__viewerE2E!.counts())
}
