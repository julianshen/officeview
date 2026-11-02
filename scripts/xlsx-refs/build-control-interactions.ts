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
import {
  CONTROL_INTERACTION_CASES, CONTROL_SEED_HEADERS, CONTROL_SEED_ROWS,
  type ControlInteractionCase,
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
  if (c.status !== 'built') {
    console.log(`Skipping non-built case: ${c.case} (${c.status})`)
    return
  }
  const outFile = resolve(ROOT, c.file)
  const workFile = join(PIVOT_DIR, `work-${c.case}-in.xlsx`)
  const workOut = join(PIVOT_DIR, `work-${c.case}-out.xlsx`)
  if (existsSync(workFile)) rmSync(workFile)
  if (existsSync(workOut)) rmSync(workOut)
  console.log(`Building ${c.case}...`)

  const baseFile = c.base === 'controls-seed'
    ? SEED_FILE
    : resolve(ROOT, CONTROL_INTERACTION_CASES.find(x => x.case === c.base)!.file)

  const zip = await JSZip.loadAsync(readFileSync(baseFile))

  // Apply XML edits according to action
  const actions = Array.isArray(c.action) ? c.action : [c.action]
  for (const op of actions) {
    if (op.kind === 'slicerSelect') {
      const part = zip.file('xl/slicerCaches/slicerCache1.xml')
      if (part) {
        let xml = await part.async('string')
        const items = op.itemIds === 'all'
          ? '<i x="0" s="1"/><i x="1" s="1"/><i x="2" s="1"/><i x="3" nd="1"/>'
          : [0, 1, 2].map(x => `<i x="${x}"${op.itemIds.includes(x) ? ' s="1"' : ''}/>`).join('') + '<i x="3" nd="1"/>'
        xml = xml.replace(/<items count="[^"]*">.*?<\/items>/s, `<items count="4">${items}</items>`)
        zip.file('xl/slicerCaches/slicerCache1.xml', xml)
      }
    } else if (op.kind === 'slicerClear') {
      const part = zip.file('xl/slicerCaches/slicerCache1.xml')
      if (part) {
        let xml = await part.async('string')
        const items = '<i x="0" s="1"/><i x="1" s="1"/><i x="2" s="1"/><i x="3" s="1" nd="1"/>'
        xml = xml.replace(/<items count="[^"]*">.*?<\/items>/s, `<items count="4">${items}</items>`)
        zip.file('xl/slicerCaches/slicerCache1.xml', xml)
      }
    } else if (op.kind === 'timelineSelect') {
      const part = zip.file('xl/timelineCaches/timelineCache1.xml')
      if (part) {
        let xml = await part.async('string')
        const state = `<state minimalRefreshVersion="6" lastRefreshVersion="6" pivotCacheId="200334102" filterType="dateBetween"><selection startDate="${op.startDate}" endDate="${op.endDate}"/><bounds startDate="2023-01-01T00:00:00" endDate="2025-01-01T00:00:00"/></state>`
        xml = xml.replace(/<state\b[^>]*>.*?<\/state>/s, state)
        zip.file('xl/timelineCaches/timelineCache1.xml', xml)
      }
    } else if (op.kind === 'timelineClear') {
      const part = zip.file('xl/timelineCaches/timelineCache1.xml')
      if (part) {
        let xml = await part.async('string')
        const state = `<state minimalRefreshVersion="6" lastRefreshVersion="6" pivotCacheId="200334102" filterType="unknown"><bounds startDate="2023-01-01T00:00:00" endDate="2025-01-01T00:00:00"/></state>`
        xml = xml.replace(/<state\b[^>]*>.*?<\/state>/s, state)
        zip.file('xl/timelineCaches/timelineCache1.xml', xml)
      }
    } else if (op.kind === 'timelineLevel') {
      const part = zip.file('xl/timelines/timeline1.xml')
      if (part) {
        let xml = await part.async('string')
        const levelIdx = ({ years: 0, quarters: 1, months: 2, days: 3 } as const)[op.level]
        xml = xml.replace(/level="[^"]*"/, `level="${levelIdx}"`)
        xml = xml.replace(/selectionLevel="[^"]*"/, `selectionLevel="${levelIdx}"`)
        zip.file('xl/timelines/timeline1.xml', xml)
      }
    }
  }

  // Special cases:
  if (c.case === 's5-slicer-crossfilter-order') {
    const part = zip.file('xl/slicerCaches/slicerCache1.xml')
    if (part) {
      let xml = await part.async('string')
      // East has Apples (0) and Bananas (1). Cherries (2) has nd="1"
      const items = '<i x="0" s="1"/><i x="1" s="1"/><i x="2" nd="1"/><i x="3" nd="1"/>'
      xml = xml.replace(/<items count="[^"]*">.*?<\/items>/s, `<items count="4">${items}</items>`)
      zip.file('xl/slicerCaches/slicerCache1.xml', xml)
    }
    // In pivot table hide West and South
    const ptPart = zip.file('xl/pivotTables/pivotTable1.xml')
    if (ptPart) {
      let ptXml = await ptPart.async('string')
      ptXml = ptXml.replace(/<item x="1"\/>/, '<item h="1" x="1"/>')
      ptXml = ptXml.replace(/<item x="2"\/>/, '<item h="1" x="2"/>')
      zip.file('xl/pivotTables/pivotTable1.xml', ptXml)
    }
  }

  writeFileSync(workFile, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))

  // Run in Excel session
  session(workFile, workOut, [], true)
  copyFileSync(workOut, outFile)

  if (existsSync(workFile)) rmSync(workFile)
  if (existsSync(workOut)) rmSync(workOut)

  // Verify markers
  if (c.markers?.length) {
    const savedZip = await JSZip.loadAsync(readFileSync(outFile))
    let allXml = ''
    for (const [, f] of Object.entries(savedZip.files)) {
      if (f.name.endsWith('.xml')) {
        allXml += await f.async('string')
      }
    }
    for (const m of c.markers) {
      if (!allXml.includes(m)) {
        throw new Error(`Case ${c.case} lacks expected marker: "${m}"`)
      }
    }
  }

  console.log(`Built ${c.case} successfully.`)
}

export async function writeManifest(): Promise<void> {
  console.log('Writing manifest.json...')
  const entries: any[] = []
  for (const c of CONTROL_INTERACTION_CASES) {
    const filePath = resolve(ROOT, c.file)
    let actualCells: Record<string, string | number | null> | undefined
    if (existsSync(filePath)) {
      try {
        actualCells = await extractCells(filePath)
      } catch {
        // ignore
      }
    }
    entries.push({
      case: c.case,
      file: c.file,
      base: c.base,
      target: c.target,
      initialState: c.initialState,
      action: c.action,
      route: c.route,
      status: c.status,
      expected: {
        ...c.expected,
        ...(actualCells ? { actualSavedCells: actualCells } : {}),
      },
      markers: c.markers,
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
