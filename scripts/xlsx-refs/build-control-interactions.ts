/**
 * Stage 5 Excel interaction oracle builder for slicers and timelines (Task 39).
 *
 * Drives Excel 16.114 on macOS to generate:
 * - tests/fixtures/xlsx-pivots/controls/controls-seed.xlsx
 * - tests/fixtures/xlsx-pivots/controls/s5-*.xlsx
 * - tests/fixtures/xlsx-pivots/controls/manifest.json
 *
 * Run:
 *   bun scripts/xlsx-refs/build-control-interactions.ts
 */
import { copyFileSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { basename, dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { closeWorkbook, openWorkbook, q, runScript, saveAs, waitFor, workbookNames } from './excel'
import { clearSlicerItems, hidePivotFieldItems, pivotFieldXmlScope, replaceXmlExactlyOnce, selectSlicerItems } from './control-interaction-xml'
import {
  CONTROL_INTERACTION_CASES,
  type ControlInteractionCase,
  type TimelineLevel,
} from './control-interaction-cases'

const __filename = fileURLToPath(import.meta.url)
const __dirname = dirname(__filename)
const ROOT = resolve(__dirname, '../..')
const PIVOT_DIR = resolve(ROOT, 'tests/fixtures/xlsx-pivots')
const CONTROLS_DIR = resolve(PIVOT_DIR, 'controls')
const RECIPES_DIR = resolve(PIVOT_DIR, 'recipes')
const BASE_TEMPLATE = join(RECIPES_DIR, 'r1-slicer-timeline.xlsx')

const SEED_FILE = join(CONTROLS_DIR, 'controls-seed.xlsx')
const MANIFEST_FILE = join(CONTROLS_DIR, 'manifest.json')

async function editPart(zip: JSZip, partName: string, edit: (xml: string) => string): Promise<void> {
  const part = zip.file(partName)
  if (!part) throw new Error(`Missing XML part ${partName}`)
  zip.file(partName, edit(await part.async('string')))
}

async function verifyPartMarkers(zip: JSZip, c: ControlInteractionCase, markers: ControlInteractionCase['scopedMarkers'] | ControlInteractionCase['inputMarkers'], label: string): Promise<void> {
  for (const marker of markers ?? []) {
    const part = zip.file(marker.part)
    if (!part) throw new Error(`${c.case} ${label} is missing marker part ${marker.part}`)
    const partXml = await part.async('string')
    let xml = partXml
    if (marker.within) {
      xml = pivotFieldXmlScope(partXml, marker.within.value)
    }
    if (marker.match) {
      const matches = xml.split(marker.match).length - 1
      if (matches !== 1) throw new Error(`${c.case} ${label} expected one '${marker.match}' in ${marker.part}, found ${matches}`)
    }
    if (marker.absent && xml.includes(marker.absent)) throw new Error(`${c.case} ${label} found forbidden '${marker.absent}' in ${marker.part}`)
  }
}

function cellPosition(cell: string): { row: number; col: number } {
  const match = /^([A-Z]+)(\d+)$/.exec(cell)
  if (!match) throw new Error(`Invalid A1 cell ${cell}`)
  let col = 0
  for (const char of match[1]!) col = col * 26 + char.charCodeAt(0) - 64
  return { row: Number(match[2]), col }
}

async function completePivotCellMap(zip: JSZip, pivotName: string, cells: Record<string, string | number | null>): Promise<Record<string, string | number | null>> {
  const paths = Object.keys(zip.files).filter(path => /^xl\/pivotTables\/pivotTable[^/]+\.xml$/.test(path))
  const matches: string[] = []
  for (const path of paths) {
    const xml = await zip.file(path)!.async('string')
    const name = /<pivotTableDefinition\b[^>]*\bname="([^"]+)"/.exec(xml)?.[1]
    const ref = /<location\b[^>]*\bref="([A-Z]+\d+:[A-Z]+\d+)"/.exec(xml)?.[1]
    if (name === pivotName && ref) matches.push(ref)
  }
  if (matches.length !== 1) throw new Error(`Expected one pivot ${pivotName}, found ${matches.length}`)
  const [start, end] = matches[0]!.split(':').map(cellPosition)
  return Object.fromEntries(Object.entries(cells).filter(([cell]) => {
    const position = cellPosition(cell)
    return position.row >= start!.row && position.row <= end!.row && position.col >= start!.col && position.col <= end!.col
  }))
}

/** True only for an Excel-saved output that still satisfies its full case oracle. */
export async function isCaseOutputVerified(c: ControlInteractionCase, file: string): Promise<boolean> {
  if (!existsSync(file)) return false
  try {
    const zip = await JSZip.loadAsync(readFileSync(file))
    await verifyPartMarkers(zip, c, c.scopedMarkers, 'saved output')
    const cells = await extractCells(file)
    for (const pivot of c.expected.pivots ?? []) {
      const actual = await completePivotCellMap(zip, pivot.pivotName, cells)
      if (JSON.stringify(actual) !== JSON.stringify(pivot.cells)) return false
      for (const coord of pivot.absentCells ?? []) if (cells[coord] !== undefined) return false
    }
    for (const expected of c.expected.slicerCache ? [c.expected.slicerCache] : []) {
      const xml = await zip.file(expected.part ?? 'xl/slicerCaches/slicerCache1.xml')?.async('string')
      if (!xml) return false
      const block = /<items\b[^>]*>([\s\S]*?)<\/items>/.exec(xml)?.[1]
      if (block === undefined) return false
      const items = [...block.matchAll(/<i\b([^>]*)\/>/g)].map(match => {
        const attrs = match[1] ?? ''
        const x = /\bx="(\d+)"/.exec(attrs)?.[1]
        return x === undefined ? null : { x: Number(x), selected: /\bs="1"/.test(attrs), noData: /\bnd="1"/.test(attrs) }
      })
      const wanted = expected.items.map(item => ({ x: item.x, selected: item.selected, noData: item.noData ?? false }))
      if (JSON.stringify(items) !== JSON.stringify(wanted)) return false
    }
    if (c.expected.secondarySlicerCache) {
      const expected = c.expected.secondarySlicerCache
      const xml = await zip.file(expected.part ?? 'xl/slicerCaches/slicerCache1.xml')?.async('string')
      const block = xml && /<items\b[^>]*>([\s\S]*?)<\/items>/.exec(xml)?.[1]
      if (block === undefined) return false
      const items = [...block.matchAll(/<i\b([^>]*)\/>/g)].map(match => {
        const attrs = match[1] ?? ''
        const x = /\bx="(\d+)"/.exec(attrs)?.[1]
        return x === undefined ? null : { x: Number(x), selected: /\bs="1"/.test(attrs), noData: /\bnd="1"/.test(attrs) }
      })
      const wanted = expected.items.map(item => ({ x: item.x, selected: item.selected, noData: item.noData ?? false }))
      if (JSON.stringify(items) !== JSON.stringify(wanted)) return false
    }
    if (c.expected.timelineCache) {
      const xml = await zip.file('xl/timelineCaches/timelineCache1.xml')?.async('string')
      if (!xml) return false
      const cache = c.expected.timelineCache
      if (cache.filterType && !xml.includes(`filterType="${cache.filterType}"`)) return false
      if (cache.selection && (!xml.includes(`startDate="${cache.selection.startDate}"`) || !xml.includes(`endDate="${cache.selection.endDate}"`))) return false
    }
    if (c.expected.tableRows) {
      const sheet = await zip.file('xl/worksheets/sheet2.xml')?.async('string')
      if (!sheet) return false
      for (let index = 1; index <= 5; index++) {
        const rowNumber = index + 1
        const row = new RegExp(`<row\\b(?=[^>]*\\br="${rowNumber}")[^>]*>`).exec(sheet)?.[0]
        if (!row) return false
        const shouldBeVisible = c.expected.tableRows.visibleRowIndices.includes(index)
        if (shouldBeVisible === /\bhidden="1"/.test(row)) return false
      }
    }
    return true
  } catch {
    return false
  }
}

export async function controlsSeedVerified(file: string): Promise<boolean> {
  if (!existsSync(file)) return false
  try {
    const cells = await extractCells(file)
    const data = await extractCells(file, 'xl/worksheets/sheet2.xml')
    return cells.A3 === '加總 - Sales' && cells.A4 === '列標籤' && cells.B4 === 'Apples' && cells.C4 === 'Bananas' && cells.D4 === 'Cherries' && cells.E4 === '總計' && cells.E8 === 150
      && data.A2 === 45291 && data.B2 === 'East' && data.C2 === 'Apples' && data.D2 === 10
      && data.A6 === 45383 && data.B6 === 'South' && data.C6 === 'Apples' && data.D6 === 50
  } catch {
    return false
  }
}

function assertControlTarget(c: ControlInteractionCase, operationName: string): void {
  if (operationName !== c.target.controlName) {
    throw new Error(`${c.case}: action targets ${operationName}; expected ${c.target.controlName}`)
  }
}

function closeAny(names: string[]): void {
  try {
    for (const open of workbookNames() ?? []) {
      if (names.includes(open)) closeWorkbook(open, false)
    }
  } catch {
    // ignore
  }
}

function session(input: string, saved: string, lines: string[] = [], refreshPivot = true): void {
  const name = openWorkbook(input)
  try {
    const script = [...lines]
    if (refreshPivot) {
      script.push(`set p to pivot table 1 of worksheet "工作表1" of workbook "${q(name)}"`, `refresh table p`)
    }
    if (script.length) {
      runScript(script, 180_000)
    }
    if (existsSync(saved)) rmSync(saved)
    saveAs(name, saved)
    waitFor(() => existsSync(saved), `write ${basename(saved)}`)
    closeAny([name, basename(saved)])
  } catch (e) {
    closeAny([name, basename(saved)])
    throw e
  }
}

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

export async function buildControlsSeed(): Promise<void> {
  if (await controlsSeedVerified(SEED_FILE)) {
    console.log('Skipping controls-seed.xlsx (existing Excel output verifies).')
    return
  }
  console.log('Building controls-seed.xlsx...')
  mkdirSync(CONTROLS_DIR, { recursive: true })
  const step1 = join(PIVOT_DIR, 'work-seed-step1.xlsx')
  const step2 = join(PIVOT_DIR, 'work-seed-step2.xlsx')
  const workSeed = join(PIVOT_DIR, 'work-controls-seed.xlsx')
  if (existsSync(step1)) rmSync(step1)
  if (existsSync(step2)) rmSync(step2)
  if (existsSync(workSeed)) rmSync(workSeed)
  copyFileSync(BASE_TEMPLATE, step1)

  // Step 1: Write Data sheet using AppleScript
  const name1 = openWorkbook(step1)
  runScript([
    `set ws to worksheet "Data" of workbook "${q(name1)}"`,
    `delete range (range "7:41" of ws)`,
    `set value of cell "A1" of ws to "Date"`,
    `set value of cell "B1" of ws to "Region"`,
    `set value of cell "C1" of ws to "Product"`,
    `set value of cell "D1" of ws to "Sales"`,
    `delete column 5 of ws`,
    `set value of cell "A2" of ws to 45291`,
    `set value of cell "B2" of ws to "East"`,
    `set value of cell "C2" of ws to "Apples"`,
    `set value of cell "D2" of ws to 10`,
    `set value of cell "A3" of ws to 45292`,
    `set value of cell "B3" of ws to "East"`,
    `set value of cell "C3" of ws to "Bananas"`,
    `set value of cell "D3" of ws to 20`,
    `set value of cell "A4" of ws to 45337`,
    `set value of cell "B4" of ws to "West"`,
    `set value of cell "C4" of ws to "Apples"`,
    `set value of cell "D4" of ws to 30`,
    `set value of cell "A5" of ws to 45382`,
    `set value of cell "B5" of ws to "West"`,
    `set value of cell "C5" of ws to "Cherries"`,
    `set value of cell "D5" of ws to 40`,
    `set value of cell "A6" of ws to 45383`,
    `set value of cell "B6" of ws to "South"`,
    `set value of cell "C6" of ws to "Apples"`,
    `set value of cell "D6" of ws to 50`,
    `set number format of range "A2:A6" of ws to "yyyy-mm-dd"`,
  ])
  if (existsSync(step2)) rmSync(step2)
  saveAs(name1, step2)
  waitFor(() => existsSync(step2), `write ${basename(step2)}`)
  closeAny([name1, basename(step2)])

  // Step 2: XML edit: update worksheetSource ref to A1:D6 and orient Product as col field
  const zip = await JSZip.loadAsync(readFileSync(step2))
  let cacheXml = await zip.file('xl/pivotCache/pivotCacheDefinition1.xml')!.async('string')
  cacheXml = cacheXml.replace(/<worksheetSource ref="[^"]*" sheet="Data"\/>/, `<worksheetSource ref="A1:D6" sheet="Data"/>`)
  zip.file('xl/pivotCache/pivotCacheDefinition1.xml', cacheXml)

  let ptXml = await zip.file('xl/pivotTables/pivotTable1.xml')!.async('string')
  ptXml = ptXml.replace(/<colItems count="1"><i\/><\/colItems>/, `<colFields count="1"><field x="1"/></colFields><colItems count="1"><i/></colItems>`)
  ptXml = ptXml.replace(/<pivotField showAll="0">/, `<pivotField axis="axisCol" showAll="0">`)
  ptXml = ptXml.replace(/ h="1"/g, '')
  zip.file('xl/pivotTables/pivotTable1.xml', ptXml)

  // SlicerCache: all items selected
  let scXml = await zip.file('xl/slicerCaches/slicerCache1.xml')!.async('string')
  scXml = scXml.replace(/<items count="[^"]*">.*?<\/items>/s, `<items count="3"><i x="0" s="1"/><i x="1" s="1"/><i x="2" s="1"/></items>`)
  zip.file('xl/slicerCaches/slicerCache1.xml', scXml)

  writeFileSync(step2, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))

  // Step 3: Open in Excel, refresh and save as SEED_FILE
  session(step2, workSeed, [], true)
  copyFileSync(workSeed, SEED_FILE)

  if (existsSync(step1)) rmSync(step1)
  if (existsSync(step2)) rmSync(step2)
  if (existsSync(workSeed)) rmSync(workSeed)
  console.log('Built controls-seed.xlsx successfully.')
}

export async function buildCase(c: ControlInteractionCase): Promise<void> {
  if (c.route === 'recipe') {
    console.log(`Skipping non-built case: ${c.case} (${c.status})`)
    return
  }
  const outFile = resolve(ROOT, c.file)
  const workFile = join(PIVOT_DIR, `work-${c.case}-in.xlsx`)
  const workOut = join(PIVOT_DIR, `work-${c.case}-out.xlsx`)
  if (await isCaseOutputVerified(c, outFile)) {
    console.log(`Skipping ${c.case} (existing Excel output verifies).`)
    return
  }
  if (existsSync(workFile) || existsSync(workOut)) {
    throw new Error(`${c.case} has preserved prior work files; inspect ${workFile} and ${workOut} before retrying`)
  }
  if (existsSync(outFile)) {
    const previous = join(PIVOT_DIR, `work-${c.case}-previous-output.xlsx`)
    if (existsSync(previous)) throw new Error(`${c.case} has an unverified output and preserved copy already exists: ${previous}`)
    copyFileSync(outFile, previous)
    rmSync(outFile)
    console.log(`Preserved prior unverified output at ${previous}`)
  }
  console.log(`Building ${c.case}...`)

  const baseFile = c.base === 'controls-seed'
    ? SEED_FILE
    : resolve(ROOT, CONTROL_INTERACTION_CASES.find(x => x.case === c.base)!.file)

  const zip = await JSZip.loadAsync(readFileSync(baseFile))

  // Apply XML edits according to action
  const actions = Array.isArray(c.action) ? c.action : [c.action]
  for (const op of actions) {
    if (op.kind === 'slicerSelect') {
      assertControlTarget(c, op.name)
      const cache = c.expected.slicerCache
      const partName = cache?.part ?? 'xl/slicerCaches/slicerCache1.xml'
      await editPart(zip, partName, xml => selectSlicerItems(xml, op.itemIds === 'all' ? Array.from({ length: Number(/<items\b[^>]*count="(\d+)"/.exec(xml)?.[1] ?? 0) }, (_, x) => x) : op.itemIds))
    } else if (op.kind === 'slicerClear') {
      assertControlTarget(c, op.name)
      const cache = c.expected.slicerCache
      if (!cache) throw new Error(`${c.case}: slicerClear requires expected cache identity`)
      await editPart(zip, cache.part ?? 'xl/slicerCaches/slicerCache1.xml', clearSlicerItems)
    } else if (op.kind === 'timelineSelect') {
      assertControlTarget(c, op.name)
      await editPart(zip, 'xl/timelineCaches/timelineCache1.xml', xml => replaceXmlExactlyOnce(xml, /<state\b[^>]*>[\s\S]*?<\/state>/, `<state minimalRefreshVersion="6" lastRefreshVersion="6" pivotCacheId="200334102" filterType="dateBetween"><selection startDate="${op.startDate}" endDate="${op.endDate}"/><bounds startDate="2023-01-01T00:00:00" endDate="2025-01-01T00:00:00"/></state>`, 'timeline selection'))
    } else if (op.kind === 'timelineClear') {
      assertControlTarget(c, op.name)
      await editPart(zip, 'xl/timelineCaches/timelineCache1.xml', xml => replaceXmlExactlyOnce(xml, /<state\b[^>]*>[\s\S]*?<\/state>/, '<state minimalRefreshVersion="6" lastRefreshVersion="6" pivotCacheId="200334102" filterType="unknown"><bounds startDate="2023-01-01T00:00:00" endDate="2025-01-01T00:00:00"/></state>', 'timeline clear'))
    } else if (op.kind === 'timelineLevel') {
      assertControlTarget(c, op.name)
      const levelIdxByName: Record<TimelineLevel, number> = { years: 0, quarters: 1, months: 2, days: 3 }
      const levelIdx = levelIdxByName[op.level as TimelineLevel]
      await editPart(zip, 'xl/timelines/timeline1.xml', xml => {
        xml = replaceXmlExactlyOnce(xml, /level="[^"]*"/, `level="${levelIdx}"`, 'timeline level')
        return replaceXmlExactlyOnce(xml, /selectionLevel="[^"]*"/, `selectionLevel="${levelIdx}"`, 'timeline selection level')
      })
    } else if (op.kind === 'pivotManualFilter') {
      if (c.target.controlName !== 'Slicer_Product' || op.fieldName !== 'Product') throw new Error(`${c.case}: manual filter target is not the expected Product field`)
      await editPart(zip, 'xl/pivotTables/pivotTable1.xml', xml => hidePivotFieldItems(xml, 'axisCol', op.hiddenItemIds))
    } else {
      throw new Error(`${c.case}: XML builder has no handler for ${op.kind}; use the owner recipe route`)
    }
  }

  await verifyPartMarkers(zip, c, c.inputMarkers, 'prepared input')

  writeFileSync(workFile, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))

  // Run in Excel session
  session(workFile, workOut, [], true)
  copyFileSync(workOut, outFile)

  // Verify markers only in the declared package part; a required match must be unique.
  const savedZip = await JSZip.loadAsync(readFileSync(outFile))
  await verifyPartMarkers(savedZip, c, c.scopedMarkers, 'saved output')
  if (!await isCaseOutputVerified(c, outFile)) throw new Error(`${c.case} Excel output failed its cell/control oracle; preserved input ${workFile} and Excel output ${workOut}`)

  if (existsSync(workFile)) rmSync(workFile)
  if (existsSync(workOut)) rmSync(workOut)

  console.log(`Built ${c.case} successfully.`)
}

export async function writeManifest(): Promise<void> {
  console.log('Writing manifest.json...')
  const entries: unknown[] = []
  for (const c of CONTROL_INTERACTION_CASES) {
    entries.push({
      case: c.case,
      file: c.file,
      base: c.base,
      target: c.target,
      initialState: c.initialState,
      action: c.action,
      route: c.route,
      status: c.status,
      expected: c.expected,
      inputMarkers: c.inputMarkers,
      scopedMarkers: c.scopedMarkers,
      recipeNotes: c.recipeNotes,
    })
  }
  writeFileSync(MANIFEST_FILE, JSON.stringify(entries, null, 2), 'utf8')
  console.log(`Manifest written with ${entries.length} cases.`)
}

async function main() {
  await buildControlsSeed()
  for (const c of CONTROL_INTERACTION_CASES) {
    await buildCase(c)
  }
  await writeManifest()
  console.log('All Stage 5 control oracle fixtures built successfully!')
}

if (process.argv[1] && /build-control-interactions\.(?:ts|js)$/.test(process.argv[1])) {
  main().catch(e => {
    console.error('Build failed:', e)
    process.exit(1)
  })
}
