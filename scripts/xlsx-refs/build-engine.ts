/**
 * Builds the Stage 2 engine oracle (pivots plan Task 15): Excel-saved workbooks for every case in engine-cases.ts.
 *
 *   bun scripts/xlsx-refs/build-engine.ts --seed            build recipes/engine-seed.xlsx from recipes/pivot-seed.xlsx
 *   bun scripts/xlsx-refs/build-engine.ts [--case id[,id…]] build the cases (all by default) into engine/<id>.xlsx
 *   bun scripts/xlsx-refs/build-engine.ts --dry-run [--seed] [--case …]
 *                                                           print routes, AppleScript and xml diffs; touch nothing
 *   bun scripts/xlsx-refs/build-engine.ts --probes          one case per unconfirmed incantation, AppleScript route only
 *                                                           (REACHED / FAILED); writes probe files under engine/.probes
 *   --no-fallback                                           with --case: skip the xml fallback (to see the AppleScript route alone)
 *
 * Stage 3 interaction oracle (pivots plan Task 24; cases in interaction-cases.ts):
 *   bun scripts/xlsx-refs/build-engine.ts --interaction [--case id[,id…]]
 *                                    build the seeds the selected cases need and are missing (getpivotdata-seed,
 *                                    refresh-seed, refresh-seed-nodata under recipes/), then the cases into
 *                                    interaction/<id>.xlsx and interaction/manifest.json (InteractionCase entries)
 *   … --interaction --seeds          (re)build every built seed, nothing else
 *   … --interaction --dry-run [--case …]   print seeds, routes, AppleScript, xml diffs and resolved deltas; touch nothing
 *   … --interaction --probes         one case per unconfirmed incantation (AppleScript only); files under interaction/.probes
 *
 * Each case: copy the seed → (xml route: optional AppleScript prep, saved, then the part edit) → open in Excel → apply
 * the case's AppleScript → `refresh table` → save as engine/<id>.xlsx → close → check the markers in the saved parts.
 * A failed route (script error or missing marker) falls back to the case's xml route, then to its owner recipe.
 * Results go to engine/manifest.json ({ case, file, features, task, pivot, route, status }); pending cases with a
 * recipe are collected into docs/validation/xlsx-pivots-recipes-stage2.md for the owner, in one batch.
 *
 * Needs an awake, unlocked display and Excel for Mac. Every file stays inside the worktree (Excel's sandbox only
 * reaches files it was granted there). Excel renames the workbook to the saved name on save-as and writes the file
 * after the script returns, so the script waits for the file and closes by the saved name.
 */
import { copyFileSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { basename, dirname, join, relative, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { closeWorkbook, openWorkbook, q, runScript, saveAs, waitFor, workbookNames } from './excel'
import {
  ENGINE_CASES, PROBES_TO_CONFIRM, SEED_CACHE_MARKERS, SEED_PIVOT_MARKERS, SEED_RECORD_MARKERS, SOURCE_REF, engineSeedStandIn, seedDataLines, seedSourceLines, xmlDiffSummary,
  type EngineCase, type EngineRoute, type PivotParts, type Target,
} from './engine-cases'
import {
  INTERACTION_CASES, INTERACTION_PROBES, SEED_BUILDS, resolveDelta, startOf, startShapeOf, type InteractionCase, type PivotDelta, type SeedBuild, type SheetMarker,
} from './interaction-cases'
import { CHART_RE, fixturePath, loadPivot, sharedStrings, sheetCells, sheetXml } from './pivot-files'

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
const PIVOT_DIR = join(ROOT, 'tests/fixtures/xlsx-pivots')
const SEED_IN = join(PIVOT_DIR, 'recipes/pivot-seed.xlsx')
export const ENGINE_SEED = join(PIVOT_DIR, 'recipes/engine-seed.xlsx')
const OUT_DIR = join(PIVOT_DIR, 'engine')
const WORK_DIR = join(OUT_DIR, '.work')
const MANIFEST = join(OUT_DIR, 'manifest.json')
const RECIPES_DOC = join(ROOT, 'docs/validation/xlsx-pivots-recipes-stage2.md')
const PIVOT_RE = /^xl\/pivotTables\/pivotTable\d+\.xml$/, CACHE_RE = /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/
const RECORDS_RE = /^xl\/pivotCache\/pivotCacheRecords\d+\.xml$/

export interface ManifestEntry {
  case: string
  file: string
  features: string[]
  task: number
  /** The pivot table's `name` (what the engine tests match `def.name` against); Excel names it in the UI locale. */
  pivot: string
  route: EngineRoute
  status: 'built' | 'pending'
  recipe?: string
  error?: string
}

const target = (name: string): Target => {
  const wb = `workbook "${q(name)}"`
  const ws = `worksheet "Pivot" of ${wb}`
  return { wb, ws, p: `pivot table 1 of ${ws}`, data: `worksheet "Data" of ${wb}` }
}

// ---- Parts ------------------------------------------------------------------------------------------------------

async function loadZip(file: string): Promise<JSZip> { return JSZip.loadAsync(readFileSync(file)) }
const partName = (zip: JSZip, re: RegExp) => Object.keys(zip.files).filter(n => re.test(n)).sort()[0]
async function readParts(file: string): Promise<PivotParts & { pivotPart: string; cachePart: string }> {
  const zip = await loadZip(file)
  const pivotPart = partName(zip, PIVOT_RE), cachePart = partName(zip, CACHE_RE)
  if (!pivotPart || !cachePart) throw new Error(`${relative(ROOT, file)} has no pivot table or cache definition`)
  return { pivot: await zip.file(pivotPart)!.async('string'), cache: await zip.file(cachePart)!.async('string'), pivotPart, cachePart }
}
async function writeEdited(input: string, output: string, edit: (parts: PivotParts) => PivotParts): Promise<void> {
  const zip = await loadZip(input)
  const parts = await readParts(input)
  const next = edit({ pivot: parts.pivot, cache: parts.cache })
  zip.file(parts.pivotPart, next.pivot)
  zip.file(parts.cachePart, next.cache)
  writeFileSync(output, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}

// ---- Excel sessions -----------------------------------------------------------------------------------------------

/** Closes whichever of `names` Excel still has open (the work name, or the saved name once save-as renamed it). */
function closeAny(names: string[]): void {
  try {
    for (const open of workbookNames() ?? []) if (names.includes(open)) closeWorkbook(open, false)
  } catch { /* Excel busy or gone: the next open retries */ }
}
/** Opens `input`, runs `lines` then `refresh table` (unless `refresh` is false), saves as `saved` and closes. */
function session(input: string, saved: string, lines: (t: Target) => string[], refresh = true): void {
  const name = openWorkbook(input)
  const t = target(name)
  try {
    const script = [...lines(t), ...(refresh ? [`refresh table ${t.p}`] : [])]
    if (script.length) runScript(script, 180_000)
    if (existsSync(saved)) rmSync(saved)
    saveAs(name, saved)
    // Excel writes the file after the script returns.
    waitFor(() => existsSync(saved), `write ${basename(saved)}`)
    closeWorkbook(basename(saved), false)
  } catch (e) {
    closeAny([name, basename(saved)])
    throw e
  }
}
const message = (e: unknown) => String((e as Error)?.message ?? e).split('\n').find(l => /error|timed out|not found|lacks|has/i.test(l))?.trim() ?? String(e)

async function verify(
  file: string,
  markers: RegExp[],
  cacheMarkers: RegExp[] = [],
  absent: RegExp[] = [],
  cacheAbsent: RegExp[] = [],
): Promise<string> {
  const parts = await readParts(file)
  for (const m of markers) if (!m.test(parts.pivot)) throw new Error(`saved pivot lacks ${m}`)
  for (const m of cacheMarkers) if (!m.test(parts.cache)) throw new Error(`saved cache lacks ${m}`)
  for (const m of absent) if (m.test(parts.pivot)) throw new Error(`saved pivot has ${m}`)
  for (const m of cacheAbsent) if (m.test(parts.cache)) throw new Error(`saved cache has ${m}`)
  return pivotName(parts.pivot)
}
const XML_ENTITIES: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }
/** The `name` attribute of a saved pivotTableDefinition, entities decoded. */
export function pivotName(pivot: string): string {
  const m = /<pivotTableDefinition\b[^>]*?\sname="([^"]*)"/.exec(pivot)
  if (!m) throw new Error('saved pivot has no name')
  return m[1]!.replace(/&(amp|lt|gt|quot|apos);|&#(\d+);|&#x([0-9a-fA-F]+);/g, (_, e, d, h) => e ? XML_ENTITIES[e]! : String.fromCodePoint(d ? Number(d) : parseInt(h, 16)))
}
/** Checks the saved records part (text of fields on no axis is inline there, not in the shared items). */
async function verifyRecords(file: string, markers: RegExp[]): Promise<void> {
  const zip = await loadZip(file)
  const part = partName(zip, RECORDS_RE)
  if (!part) throw new Error(`${relative(ROOT, file)} has no cache records`)
  const records = await zip.file(part)!.async('string')
  for (const m of markers) if (!m.test(records)) throw new Error(`saved records lack ${m}`)
}

// ---- Seed -------------------------------------------------------------------------------------------------------

export async function buildSeed(): Promise<void> {
  mkdirSync(WORK_DIR, { recursive: true })
  const work = join(WORK_DIR, 'engine-seed-in.xlsx'), staged = join(WORK_DIR, 'engine-seed-data.xlsx'), edited = join(WORK_DIR, 'engine-seed-xml.xlsx')
  copyFileSync(SEED_IN, work)
  // Step 1 writes the source; step 2 repoints the pivot. A failing `set source data` keeps the written data and
  // takes the xml route (worksheetSource ref rewritten, then refreshed by Excel).
  const name = openWorkbook(work)
  const t = target(name)
  let viaXml = false
  try {
    runScript(seedDataLines(t.data), 300_000)
    try {
      runScript(seedSourceLines(t.p), 180_000)
    } catch (e) {
      console.log(`seed: set source data failed (${message(e)}); taking the xml route`)
      viaXml = true
    }
    const saved = viaXml ? staged : ENGINE_SEED
    if (existsSync(saved)) rmSync(saved)
    saveAs(name, saved)
    waitFor(() => existsSync(saved), `write ${basename(saved)}`)
    closeWorkbook(basename(saved), false)
  } catch (e) {
    closeAny([name, basename(staged), basename(ENGINE_SEED)])
    throw e
  }
  if (!viaXml) {
    try { await verify(ENGINE_SEED, SEED_PIVOT_MARKERS, SEED_CACHE_MARKERS); await verifyRecords(ENGINE_SEED, SEED_RECORD_MARKERS); console.log(`built ${relative(ROOT, ENGINE_SEED)} (applescript)`); return }
    catch (e) { console.log(`seed: ${message(e)}; taking the xml route`); copyFileSync(ENGINE_SEED, staged) }
  }
  await writeEdited(staged, edited, p => ({ ...p, cache: p.cache.replace(/(<worksheetSource [^>]*ref=")[^"]*"/, `$1${SOURCE_REF}"`) }))
  session(edited, ENGINE_SEED, () => [])
  await verify(ENGINE_SEED, SEED_PIVOT_MARKERS, SEED_CACHE_MARKERS)
  await verifyRecords(ENGINE_SEED, SEED_RECORD_MARKERS)
  console.log(`built ${relative(ROOT, ENGINE_SEED)} (xml)`)
}

// ---- Cases ------------------------------------------------------------------------------------------------------

async function runRoute(c: EngineCase, route: 'applescript' | 'xml', out: string): Promise<void> {
  const input = join(WORK_DIR, `${c.id}-in.xlsx`), prepped = join(WORK_DIR, `${c.id}-prep.xlsx`), edited = join(WORK_DIR, `${c.id}-xml.xlsx`)
  copyFileSync(ENGINE_SEED, input)
  try {
    if (route === 'applescript') return session(input, out, c.lines ?? (() => []))
    if (!c.xmlEdit) throw new Error('no xml edit')
    let source = input
    if (c.xmlPrep) { session(input, prepped, c.xmlPrep); source = prepped }
    await writeEdited(source, edited, c.xmlEdit)
    session(edited, out, () => [])
  } finally {
    for (const f of [input, prepped, edited]) rmSync(f, { force: true })
  }
}

export async function buildCase(c: EngineCase, fallback = true): Promise<ManifestEntry> {
  const out = join(OUT_DIR, `${c.id}.xlsx`)
  const entry = (route: EngineRoute, status: ManifestEntry['status'], pivot = '', error?: string): ManifestEntry =>
    ({ case: c.id, file: relative(ROOT, out), features: c.features, task: c.task, pivot, route, status, ...(c.recipe && status === 'pending' ? { recipe: c.recipe } : {}), ...(error ? { error } : {}) })
  if (c.route === 'recipe') return entry('recipe', 'pending')
  const routes: Array<'applescript' | 'xml'> = c.route === 'xml' ? ['xml'] : ['applescript', ...(fallback && c.fallback === 'xml' ? ['xml' as const] : [])]
  const errors: string[] = []
  for (const route of routes) {
    try {
      await runRoute(c, route, out)
      const pivot = await verify(out, c.markers, c.cacheMarkers, c.absent)
      console.log(`built ${c.id} (${route})`)
      return entry(route, 'built', pivot)
    } catch (e) {
      errors.push(`${route}: ${message(e)}`)
      console.log(`${c.id}: ${route} failed: ${message(e)}`)
    }
  }
  rmSync(out, { force: true })
  return entry(c.recipe ? 'recipe' : routes[routes.length - 1]!, 'pending', undefined, errors.join('; '))
}

function loadManifest(): ManifestEntry[] {
  return existsSync(MANIFEST) ? JSON.parse(readFileSync(MANIFEST, 'utf8')) as ManifestEntry[] : []
}
function writeManifest(entries: ManifestEntry[]): void {
  const order = new Map(ENGINE_CASES.map((c, i) => [c.id, i]))
  entries.sort((a, b) => (order.get(a.case) ?? 1e9) - (order.get(b.case) ?? 1e9))
  writeFileSync(MANIFEST, `${JSON.stringify(entries, null, 2)}\n`)
}
/** One batch for the owner: every pending case, with its recipe or the failure to look at. */
function writeRecipes(entries: ManifestEntry[]): void {
  const pending = entries.filter(e => e.status === 'pending')
  if (!pending.length) { rmSync(RECIPES_DOC, { force: true }); return }
  const byId = new Map(ENGINE_CASES.map(c => [c.id, c]))
  const lines = [
    '# XLSX pivots: Stage 2 owner recipes', '',
    'Cases that neither AppleScript nor the hand-edited XML route could make (scripts/xlsx-refs/build-engine.ts).',
    'Start each from `tests/fixtures/xlsx-pivots/recipes/engine-seed.xlsx` (open a copy, never the seed itself), apply the steps,',
    'refresh the pivot (PivotTable Analyze → Refresh), and save under the name given. Engine tasks keep these cases `pending` until the files exist.', '',
  ]
  for (const e of pending) {
    const c = byId.get(e.case)
    lines.push(`## ${e.case} (Task ${e.task})`, '', `- Measures: ${c?.measures ?? e.features.join(', ')}`)
    lines.push(e.recipe ? `- Recipe: ${e.recipe}` : '- Recipe: none yet; the automated routes failed (below), so this case needs a decision.')
    if (e.error) lines.push(`- Automated attempts: ${e.error}`)
    lines.push('')
  }
  writeFileSync(RECIPES_DOC, lines.join('\n'))
}

// ---- Dry run ----------------------------------------------------------------------------------------------------

async function seedPartsForDryRun(): Promise<{ parts: PivotParts; source: string }> {
  if (existsSync(ENGINE_SEED)) return { parts: await readParts(ENGINE_SEED), source: relative(ROOT, ENGINE_SEED) }
  return { parts: engineSeedStandIn(await readParts(SEED_IN)), source: `stand-in from ${relative(ROOT, SEED_IN)} (engine-seed.xlsx not built yet)` }
}
const clip = (s: string, n = 150) => s.length > n ? `${s.slice(0, n)}…` : s

async function dryRun(cases: readonly EngineCase[], seed: boolean): Promise<number> {
  const t = target('<case>-in.xlsx')
  if (seed) {
    console.log('== seed: recipes/pivot-seed.xlsx → recipes/engine-seed.xlsx')
    const data = seedDataLines(t.data)
    for (const l of data) console.log(`  ${clip(l, 200)}`)
    for (const l of seedSourceLines(t.p)) console.log(`  ${l}`)
    console.log(`  markers: ${[...SEED_PIVOT_MARKERS, ...SEED_CACHE_MARKERS].map(String).join(' ')}`)
  }
  const { parts, source } = await seedPartsForDryRun()
  console.log(`xml diffs against: ${source}\n`)
  let failures = 0
  const routes = new Map<string, number>()
  for (const c of cases) {
    const tag = c.route === 'applescript' && c.fallback ? 'applescript→xml' : c.route
    routes.set(tag, (routes.get(tag) ?? 0) + 1)
    console.log(`== ${c.id}  [${c.group}, task ${c.task}, route ${c.route}${c.fallback ? `, fallback ${c.fallback}` : ''}${c.recipe ? ', recipe' : ''}${c.probe ? `, probe ${c.probe}` : ''}]`)
    console.log(`   features: ${c.features.join(', ')}`)
    if (c.measures) console.log(`   measures: ${c.measures}`)
    if (c.lines) for (const l of [...c.lines(t), `refresh table ${t.p}`]) console.log(`   as> ${clip(l, 220)}`)
    if (c.xmlPrep) for (const l of c.xmlPrep(t)) console.log(`   prep> ${clip(l, 220)}`)
    if (c.xmlEdit) {
      try {
        const next = c.xmlEdit({ ...parts })
        for (const [label, before, after] of [['pivot', parts.pivot, next.pivot], ['cache', parts.cache, next.cache]] as const) {
          const { added, removed } = xmlDiffSummary(before, after)
          if (!added.length && !removed.length) continue
          console.log(`   xml ${label}: +${added.length} -${removed.length}`)
          for (const a of added.slice(0, 6)) console.log(`     + ${clip(a)}`)
          if (added.length > 6) console.log(`     + … ${added.length - 6} more`)
          for (const r of removed.slice(0, 4)) console.log(`     - ${clip(r)}`)
          if (removed.length > 4) console.log(`     - … ${removed.length - 4} more`)
        }
      } catch (e) {
        failures++
        console.log(`   xml EDIT FAILED: ${(e as Error).message}`)
      }
    }
    console.log(`   markers: ${[...c.markers.map(m => `pivot ${m}`), ...(c.cacheMarkers ?? []).map(m => `cache ${m}`), ...(c.absent ?? []).map(m => `not ${m}`)].map(s => clip(s, 120)).join(' | ') || '(none: refresh only)'}`)
    if (c.recipe) console.log(`   recipe: ${c.recipe}`)
  }
  const probes = new Set(cases.map(c => c.probe).filter((p): p is string => !!p))
  console.log(`\n${cases.length} cases: ${[...routes].map(([r, n]) => `${n} ${r}`).join(', ')}; ${cases.filter(c => c.recipe).length} with an owner recipe`)
  console.log('probe first:')
  for (const p of ['set-source-data', 'write-range', ...probes]) console.log(`  ${p}: ${PROBES_TO_CONFIRM[p] ?? ''}`)
  if (failures) console.log(`${failures} xml edits FAILED`)
  return failures
}

// ---- Probes -----------------------------------------------------------------------------------------------------

/** The AppleScript route of the first case that depends on each unconfirmed incantation; nothing is recorded. */
async function probes(): Promise<void> {
  const dir = join(OUT_DIR, '.probes')
  mkdirSync(WORK_DIR, { recursive: true }); mkdirSync(dir, { recursive: true })
  const seen = new Set<string>()
  for (const c of ENGINE_CASES) {
    if (!c.probe || seen.has(c.probe) || c.route !== 'applescript') continue
    seen.add(c.probe)
    const out = join(dir, `${c.id}.xlsx`)
    let status: string
    try {
      await runRoute(c, 'applescript', out)
      await verify(out, c.markers, c.cacheMarkers, c.absent)
      status = 'REACHED'
    } catch (e) {
      status = `FAILED ${message(e)}`
    }
    console.log(`${c.probe} (${c.id}): ${status}`)
  }
  rmSync(WORK_DIR, { recursive: true, force: true })
}

// ---- Stage 3 interaction oracle (Task 24) -----------------------------------------------------------------------

const INTERACTION_DIR = join(PIVOT_DIR, 'interaction')
const INTERACTION_MANIFEST = join(INTERACTION_DIR, 'manifest.json')
const STAGE4_RECIPES_DOC = join(ROOT, 'docs/validation/xlsx-pivots-recipes-stage4.md')

/** One `interaction/manifest.json` entry: the plan's `InteractionCase` (Task 24 & Task 34 interfaces) plus where the
 * case starts (`base`), what non-UI change it made (`setup`), and how it was built. */
export interface InteractionManifestEntry {
  case: string
  file: string
  pivot: string
  task: InteractionCase['task']
  seed: InteractionCase['seed']
  delta: PivotDelta[]
  measures?: string
  /** The start file when it is not the seed's own (relative to tests/fixtures/xlsx-pivots, without .xlsx). */
  base?: string
  /** A change no PivotDelta expresses, made before `delta`: the case is a measurement or a base, not a delta oracle. */
  setup?: string
  route: InteractionCase['route']
  status: 'built' | 'pending'
  recipe?: string
  error?: string
}

async function verifySheets(file: string, markers: readonly SheetMarker[] = []): Promise<void> {
  if (!markers.length) return
  const zip = await loadZip(file)
  for (const m of markers) {
    const xml = await sheetXml(zip, m.sheet)
    if (xml === undefined) throw new Error(`saved file has no sheet ${m.sheet}`)
    if (m.absent ? m.re.test(xml) : !m.re.test(xml)) throw new Error(`saved sheet ${m.sheet} ${m.absent ? 'has' : 'lacks'} ${m.re}`)
  }
}
async function verifyRecordsAbsent(file: string, absent: readonly RegExp[]): Promise<void> {
  const zip = await loadZip(file)
  const part = partName(zip, RECORDS_RE)
  if (!part) throw new Error(`${relative(ROOT, file)} has no cache records`)
  const records = await zip.file(part)!.async('string')
  for (const m of absent) if (m.test(records)) throw new Error(`saved records have ${m}`)
}
async function verifyNoRecords(file: string): Promise<void> {
  if (partName(await loadZip(file), RECORDS_RE)) throw new Error('saved file still has a records part (saveData="0" not kept)')
}
async function verifyChart(file: string): Promise<void> {
  const zip = await loadZip(file)
  const charts = await Promise.all(Object.keys(zip.files).filter(n => CHART_RE.test(n)).map(n => zip.file(n)!.async('string')))
  if (!charts.some(c => /<c:pivotSource>/.test(c))) throw new Error('saved file has no pivot chart (c:pivotSource)')
}

/** Cache-definition attributes set and, with `dropRecords`, the records part removed with its relationship, the
 * definition's r:id and its content-type override (the refresh seeds' xml fallback). */
async function editCachePackage(input: string, output: string, attrs: Record<string, string>, dropRecords: boolean): Promise<void> {
  const zip = await loadZip(input)
  const parts = await readParts(input)
  let cache = parts.cache.replace(/<pivotCacheDefinition\b[^>]*>/, open => {
    for (const [k, v] of Object.entries(attrs)) open = new RegExp(`\\s${k}="`).test(open) ? open.replace(new RegExp(`\\s${k}="[^"]*"`), ` ${k}="${v}"`) : open.replace(/>$/, ` ${k}="${v}">`)
    return open
  })
  if (dropRecords) {
    const records = partName(zip, RECORDS_RE)
    if (records) {
      const relsPath = parts.cachePart.replace(/([^/]+)$/, '_rels/$1.rels')
      const rels = await zip.file(relsPath)?.async('string')
      const id = rels && new RegExp(`<Relationship [^>]*Id="([^"]+)"[^>]*Target="[^"]*${basename(records)}"`).exec(rels)?.[1]
      if (rels && id) {
        const kept = rels.replace(new RegExp(`<Relationship [^>]*Id="${id}"[^>]*/>`), '')
        if (/<Relationship /.test(kept)) zip.file(relsPath, kept); else zip.remove(relsPath)
        cache = cache.replace(new RegExp(`\\sr:id="${id}"`), '')
      }
      zip.remove(records)
      const types = await zip.file('[Content_Types].xml')!.async('string')
      zip.file('[Content_Types].xml', types.replace(new RegExp(`<Override [^>]*PartName="/${records.replace(/[.]/g, '\\.')}"[^>]*/>`), ''))
    }
  }
  zip.file(parts.cachePart, cache)
  writeFileSync(output, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}

async function verifySeedBuild(s: SeedBuild, out: string): Promise<void> {
  await verify(out, s.markers, s.cacheMarkers)
  if (s.recordMarkers) await verifyRecords(out, s.recordMarkers)
  if (s.recordAbsent) await verifyRecordsAbsent(out, s.recordAbsent)
  if (s.noRecords) await verifyNoRecords(out)
  await verifySheets(out, s.sheetMarkers)
}
export async function buildInteractionSeed(s: SeedBuild, fallback = true): Promise<void> {
  mkdirSync(WORK_DIR, { recursive: true })
  const out = fixturePath(s.file), input = join(WORK_DIR, `${s.id}-in.xlsx`), staged = join(WORK_DIR, `${s.id}-staged.xlsx`)
  try {
    copyFileSync(ENGINE_SEED, input)
    try {
      session(input, out, s.lines, s.refresh)
      await verifySeedBuild(s, out)
      console.log(`built ${relative(ROOT, out)} (applescript)`)
      return
    } catch (e) {
      console.log(`${s.id}: applescript failed: ${message(e)}`)
      if (!s.fallback || !fallback) { rmSync(out, { force: true }); throw e }
    }
    // The reachable part through AppleScript (the data edit, no refresh), then the cache flags in the package.
    copyFileSync(ENGINE_SEED, input)
    session(input, staged, s.fallback.lines, false)
    await editCachePackage(staged, out, s.fallback.cache, s.fallback.dropRecords ?? false)
    await verifySeedBuild(s, out)
    console.log(`built ${relative(ROOT, out)} (xml fallback: Excel has not opened this file; the first case that starts from it does)`)
  } finally {
    for (const f of [input, staged]) rmSync(f, { force: true })
  }
}

async function pivotSheetXml(zip: JSZip, pivotPartName: string): Promise<string> {
  const pivotBase = basename(pivotPartName)
  for (const name of Object.keys(zip.files)) {
    if (/^xl\/worksheets\/_rels\/.*\.rels$/.test(name)) {
      const rels = await zip.file(name)!.async('string')
      if (rels.includes(pivotBase)) {
        const sheetPartName = name.replace('/_rels/', '/').replace(/\.rels$/, '')
        const sheet = await zip.file(sheetPartName)?.async('string')
        if (sheet) return sheet
      }
    }
  }
  const firstSheet = Object.keys(zip.files).find(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))
  return (firstSheet ? await zip.file(firstSheet)?.async('string') : '') ?? ''
}

export async function readPivotLocationCells(file: string): Promise<{ loc: string; cells: Map<string, string> }> {
  const zip = await loadZip(file)
  const strings = await sharedStrings(zip)
  const pName = Object.keys(zip.files).find(n => PIVOT_RE.test(n))
  if (!pName) throw new Error(`${file}: no pivotTable part`)
  const pXml = await zip.file(pName)!.async('string')
  const loc = /<location\b[^>]*\bref="([^"]*)"/.exec(pXml)?.[1]
  if (!loc) throw new Error(`${file}: pivot has no location ref`)
  const [tl, br = tl] = loc.split(':') as [string, string?]
  const toCoord = (s: string) => {
    const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(s) ?? []
    const col = [...letters!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1
    const row = Number(digits) - 1
    return { row, col }
  }
  const p1 = toCoord(tl), p2 = toCoord(br)
  const r1 = Math.min(p1.row, p2.row), c1 = Math.min(p1.col, p2.col), r2 = Math.max(p1.row, p2.row), c2 = Math.max(p1.col, p2.col)
  const sXml = await pivotSheetXml(zip, pName)
  const inLoc = sheetCells(sXml, strings).filter(c => c.row >= r1 && c.row <= r2 && c.col >= c1 && c.col <= c2)
  return { loc, cells: new Map(inLoc.map(c => [c.ref, c.value ?? ''])) }
}

export async function verifyXmlRecomputation(edited: string, out: string): Promise<{ cellDiffs: number; locChanged: boolean }> {
  const [editedPivot, outPivot] = await Promise.all([readPivotLocationCells(edited), readPivotLocationCells(out)])
  const locChanged = editedPivot.loc !== outPivot.loc
  const allRefs = new Set([...editedPivot.cells.keys(), ...outPivot.cells.keys()])
  let cellDiffs = 0
  for (const ref of allRefs) {
    if (editedPivot.cells.get(ref) !== outPivot.cells.get(ref)) cellDiffs++
  }
  if (cellDiffs === 0) {
    throw new Error(`XML route: saved pivot cells in location range (${outPivot.loc}) are unchanged from edited input (no recomputation)`)
  }
  return { cellDiffs, locChanged }
}

async function runInteraction(c: InteractionCase, route: 'applescript' | 'xml', out: string): Promise<void> {
  const start = fixturePath(startOf(c))
  if (!existsSync(start)) throw new Error(`start file ${startOf(c)}.xlsx is not built`)
  const input = join(WORK_DIR, `${c.id}-in.xlsx`), prepped = join(WORK_DIR, `${c.id}-prep.xlsx`), edited = join(WORK_DIR, `${c.id}-xml.xlsx`)
  copyFileSync(start, input)
  try {
    if (route === 'applescript') return session(input, out, c.lines ?? (() => []))
    if (!c.xmlEdit) throw new Error('no xml edit')
    let source = input
    if (c.xmlPrep) { session(input, prepped, c.xmlPrep); source = prepped }
    await writeEdited(source, edited, c.xmlEdit)
    session(edited, out, c.post ?? (() => []))
    await verifyXmlRecomputation(edited, out)
  } finally {
    for (const f of [input, prepped, edited]) rmSync(f, { force: true })
  }
}
export async function verifyInteraction(c: InteractionCase, out: string): Promise<string> {
  const startFile = fixturePath(startOf(c))
  if (existsSync(startFile)) {
    const startParts = await readParts(startFile)
    const outParts = await readParts(out)
    if (c.delta.length > 0 && startParts.pivot === outParts.pivot && startParts.cache === outParts.cache) {
      throw new Error(`saved pivot is identical to start file ${startOf(c)}: no-op interaction`)
    }
    if (c.route === 'xml') {
      const [startPivot, outPivot] = await Promise.all([readPivotLocationCells(startFile), readPivotLocationCells(out)])
      const allRefs = new Set([...startPivot.cells.keys(), ...outPivot.cells.keys()])
      let cellDiffs = 0
      for (const ref of allRefs) {
        if (startPivot.cells.get(ref) !== outPivot.cells.get(ref)) cellDiffs++
      }
      if (cellDiffs === 0) {
        throw new Error(`XML route: saved pivot cells in location range (${outPivot.loc}) are unchanged from start ${startOf(c)} (no recomputation)`)
      }
    }
  }
  const pivot = await verify(out, c.markers, c.cacheMarkers, c.absent, c.cacheAbsent)
  if (c.recordMarkers) await verifyRecords(out, c.recordMarkers)
  await verifySheets(out, c.sheetMarkers)
  if (c.chart) await verifyChart(out)
  return pivot
}

/** The case's deltas in engine indices, resolved against its start file; a filter delta takes the filter Excel saved
 * for that field in the case file when there is exactly one (what Excel applied, `filterVal` and all). */
async function manifestDelta(c: InteractionCase, out?: string): Promise<PivotDelta[]> {
  const shape = startShapeOf(await readParts(fixturePath(startOf(c))))
  const delta = c.delta.map(spec => resolveDelta(spec, shape))
  if (!out || !delta.some(d => d.kind === 'filter')) return delta
  try {
    const { def } = await loadPivot(out)
    return delta.map(d => {
      if (d.kind !== 'filter') return d
      const saved = def.filters.filter(f => f.fld === d.field)
      return saved.length === 1 ? { ...d, filter: saved[0]! } : d
    })
  } catch (e) {
    console.log(`${c.id}: kept the literal filter (${message(e)})`)
    return delta
  }
}

export async function buildInteractionCase(c: InteractionCase, fallback = true): Promise<InteractionManifestEntry> {
  const out = join(INTERACTION_DIR, `${c.id}.xlsx`)
  const entry = async (route: InteractionManifestEntry['route'], status: InteractionManifestEntry['status'], pivot = '', error?: string): Promise<InteractionManifestEntry> => {
    let delta: PivotDelta[] = []
    try { delta = existsSync(fixturePath(startOf(c))) ? await manifestDelta(c, status === 'built' ? out : undefined) : [] }
    catch (e) { error = [error, `delta: ${message(e)}`].filter(Boolean).join('; '); status = 'pending' }
    return {
      case: c.id, file: relative(ROOT, out), pivot, task: c.task, seed: c.seed, delta, measures: c.measures,
      ...(c.base ? { base: c.base } : {}), ...(c.setup ? { setup: c.setup } : {}), route, status,
      ...(c.recipe && status === 'pending' ? { recipe: c.recipe } : {}),
      ...(error ? { error } : {}),
    }
  }
  if (c.route === 'recipe') {
    if (existsSync(out)) {
      try {
        const pivot = await verifyInteraction(c, out)
        console.log(`verified existing recipe save ${c.id}`)
        return entry('recipe', 'built', pivot)
      } catch (e) {
        return entry('recipe', 'pending', undefined, `recipe verify: ${message(e)}`)
      }
    }
    return entry('recipe', 'pending')
  }
  const routes: Array<'applescript' | 'xml'> = c.route === 'xml' ? ['xml'] : ['applescript', ...(fallback && c.fallback === 'xml' ? ['xml' as const] : [])]
  const errors: string[] = []
  for (const route of routes) {
    try {
      await runInteraction(c, route, out)
      const pivot = await verifyInteraction(c, out)
      console.log(`built ${c.id} (${route})`)
      return entry(route, 'built', pivot)
    } catch (e) {
      errors.push(`${route}: ${message(e)}`)
      console.log(`${c.id}: ${route} failed: ${message(e)}`)
      if (/is not built/.test(message(e))) break
    }
  }
  rmSync(out, { force: true })
  return entry(c.recipe ? 'recipe' : routes[routes.length - 1]!, 'pending', undefined, errors.join('; '))
}

function loadInteractionManifest(): InteractionManifestEntry[] {
  return existsSync(INTERACTION_MANIFEST) ? JSON.parse(readFileSync(INTERACTION_MANIFEST, 'utf8')) as InteractionManifestEntry[] : []
}
function writeInteractionManifest(entries: InteractionManifestEntry[]): void {
  const order = new Map(INTERACTION_CASES.map((c, i) => [c.id, i]))
  entries.sort((a, b) => (order.get(a.case) ?? 1e9) - (order.get(b.case) ?? 1e9))
  writeFileSync(INTERACTION_MANIFEST, `${JSON.stringify(entries, null, 2)}\n`)
}
/** Single Stage 4 batch recipe doc (spec §11): written when any Stage 4 case (task 35..37) remains pending. */
export function writeInteractionRecipes(entries: readonly InteractionManifestEntry[]): void {
  const pending = entries.filter(e => e.status === 'pending' && (e.task === 35 || e.task === 36 || e.task === 37))
  if (!pending.length) { rmSync(STAGE4_RECIPES_DOC, { force: true }); return }
  const byId = new Map(INTERACTION_CASES.map(c => [c.id, c]))
  const lines = [
    '# XLSX pivots: Stage 4 interaction recipes', '',
    'Cases that neither AppleScript nor a verified Excel-refreshed XML edit could complete (`scripts/xlsx-refs/build-engine.ts --interaction`).',
    'Open a copy of each case\'s start workbook (never overwrite `recipes/*` or `engine/*`), apply the exact Excel UI clicks,',
    'refresh the pivot (`PivotTable Analyze → Refresh`), and save to the output path listed.', '',
  ]
  for (const e of pending) {
    const c = byId.get(e.case)
    const start = c ? `tests/fixtures/xlsx-pivots/${startOf(c)}.xlsx` : `tests/fixtures/xlsx-pivots/recipes/${e.seed}.xlsx`
    lines.push(`## ${e.case} (Task ${e.task})`, '')
    lines.push(`- Start file: \`${start}\``)
    lines.push(`- Output file: \`${e.file}\``)
    if (e.measures) lines.push(`- Measures: ${e.measures}`)
    lines.push(e.recipe || c?.recipe ? `- Recipe: ${e.recipe ?? c?.recipe}` : '- Recipe: none yet; automated routes failed (below) and require System Events UI recipe or owner save.')
    if (e.error) lines.push(`- Automated attempts: ${e.error}`)
    lines.push('')
  }
  writeFileSync(STAGE4_RECIPES_DOC, lines.join('\n'))
}

/** The parts the dry run diffs a case against: its start file when built, else the nearest built ancestor (a base
 * that is itself an interaction case falls back to its own start), else the engine seed stand-in. */
async function startPartsForDryRun(c: InteractionCase): Promise<{ parts: PivotParts; source: string; exact: boolean }> {
  let rel: string | undefined = startOf(c)
  for (let guard = 0; rel && guard < 8; guard++) {
    if (existsSync(fixturePath(rel))) return { parts: await readParts(fixturePath(rel)), source: rel, exact: rel === startOf(c) }
    const parent: InteractionCase | undefined = rel.startsWith('interaction/') ? INTERACTION_CASES.find(x => `interaction/${x.id}` === rel) : undefined
    rel = parent ? startOf(parent) : undefined
  }
  const seed = await seedPartsForDryRun()
  return { parts: seed.parts, source: seed.source, exact: false }
}

export async function interactionDryRun(cases: readonly InteractionCase[], seeds: readonly SeedBuild[]): Promise<number> {
  const t = target('<case>-in.xlsx')
  let failures = 0
  for (const s of seeds) {
    console.log(`== seed ${s.id}  [engine-seed → ${s.file}.xlsx${s.refresh ? '' : ', saved without refresh'}${existsSync(fixturePath(s.file)) ? ', built' : ', not built'}, probe ${s.probe}]`)
    console.log(`   measures: ${s.measures}`)
    for (const l of s.lines(t)) console.log(`   as> ${clip(l, 220)}`)
    if (s.fallback) console.log(`   fallback: ${s.fallback.lines(t).length} data lines, then cache ${JSON.stringify(s.fallback.cache)}${s.fallback.dropRecords ? ' and the records part dropped' : ''}`)
    console.log(`   markers: ${[...s.markers.map(m => `pivot ${m}`), ...(s.cacheMarkers ?? []).map(m => `cache ${m}`), ...(s.recordAbsent ?? []).map(m => `records not ${m}`),
      ...(s.noRecords ? ['no records part'] : []), ...(s.sheetMarkers ?? []).map(m => `${m.sheet} ${m.absent ? 'not ' : ''}${m.re}`)].map(x => clip(x, 120)).join(' | ')}`)
  }
  const routes = new Map<string, number>()
  for (const c of cases) {
    const tag = c.route === 'applescript' && c.fallback ? 'applescript→xml' : c.route
    routes.set(tag, (routes.get(tag) ?? 0) + 1)
    console.log(`== ${c.id}  [task ${c.task}, seed ${c.seed}${c.base ? `, base ${c.base}` : ''}, route ${c.route}${c.fallback ? `, fallback ${c.fallback}` : ''}${c.probe ? `, probe ${c.probe}` : ''}]`)
    if (c.setup) console.log(`   setup: ${c.setup}`)
    console.log(`   measures: ${c.measures}`)
    if (c.lines) for (const l of [...c.lines(t), `refresh table ${t.p}`]) console.log(`   as> ${clip(l, 220)}`)
    if (c.xmlPrep) for (const l of c.xmlPrep(t)) console.log(`   prep> ${clip(l, 220)}`)
    if (c.post) for (const l of c.post(t)) console.log(`   post> ${clip(l, 220)}`)
    const { parts, source, exact } = await startPartsForDryRun(c)
    if (c.xmlEdit) {
      try {
        const next = c.xmlEdit({ ...parts })
        console.log(`   xml diff against ${source}${exact ? '' : ' (stand-in: the start file is not built yet)'}`)
        for (const [label, before, after] of [['pivot', parts.pivot, next.pivot], ['cache', parts.cache, next.cache]] as const) {
          const { added, removed } = xmlDiffSummary(before, after)
          if (!added.length && !removed.length) continue
          console.log(`   xml ${label}: +${added.length} -${removed.length}`)
          for (const a of added.slice(0, 6)) console.log(`     + ${clip(a)}`)
          if (added.length > 6) console.log(`     + … ${added.length - 6} more`)
          for (const r of removed.slice(0, 4)) console.log(`     - ${clip(r)}`)
          if (removed.length > 4) console.log(`     - … ${removed.length - 4} more`)
        }
      } catch (e) {
        if (exact) failures++
        console.log(`   xml EDIT FAILED${exact ? '' : ' (on a stand-in)'}: ${(e as Error).message}`)
      }
    }
    console.log(`   markers: ${[...c.markers.map(m => `pivot ${m}`), ...(c.cacheMarkers ?? []).map(m => `cache ${m}`), ...(c.absent ?? []).map(m => `not ${m}`),
      ...(c.cacheAbsent ?? []).map(m => `cache not ${m}`), ...(c.recordMarkers ?? []).map(m => `records ${m}`), ...(c.sheetMarkers ?? []).map(m => `${m.sheet} ${m.re}`), ...(c.chart ? ['pivot chart'] : [])].map(s => clip(s, 120)).join(' | ')}`)
    if (c.recipe) console.log(`   recipe: ${c.recipe}`)
    if (exact) {
      try { console.log(`   delta: ${JSON.stringify(c.delta.map(spec => resolveDelta(spec, startShapeOf(parts))).map(d => d.kind === 'filter' ? { ...d, filter: `${d.filter.type} (literal; Excel's saved filter replaces it)` } : d))}`) }
      catch (e) { failures++; console.log(`   delta RESOLUTION FAILED: ${(e as Error).message}`) }
    } else console.log(`   delta (by name; resolved once ${startOf(c)} is built): ${JSON.stringify(c.delta)}`)
  }
  console.log(`\n${seeds.length} seeds, ${cases.length} cases: ${[...routes].map(([r, n]) => `${n} ${r}`).join(', ')}; ` +
    `${cases.filter(c => c.setup).length} with a setup, ${cases.filter(c => c.base).length} with a base`)
  console.log('probe first:')
  for (const p of new Set([...seeds.map(s => s.probe), ...cases.map(c => c.probe).filter((p): p is string => !!p)])) console.log(`  ${p}: ${INTERACTION_PROBES[p] ?? ''}`)
  if (failures) console.log(`${failures} xml edits or deltas FAILED`)
  return failures
}

/** The first AppleScript case per unconfirmed incantation whose start file exists; nothing is recorded. */
async function interactionProbes(selected: readonly InteractionCase[] = INTERACTION_CASES): Promise<void> {
  const dir = join(INTERACTION_DIR, '.probes')
  mkdirSync(WORK_DIR, { recursive: true }); mkdirSync(dir, { recursive: true })
  const seen = new Set<string>()
  for (const c of selected) {
    if (!c.probe || seen.has(c.probe) || c.route !== 'applescript' || !existsSync(fixturePath(startOf(c)))) continue
    seen.add(c.probe)
    let status: string
    try { await runInteraction(c, 'applescript', join(dir, `${c.id}.xlsx`)); await verifyInteraction(c, join(dir, `${c.id}.xlsx`)); status = 'REACHED' }
    catch (e) { status = `FAILED ${message(e)}` }
    console.log(`${c.probe} (${c.id}): ${status}`)
    // Only this run's probe file: the directory also holds a tracked Stage 3 probe save.
    rmSync(join(dir, `${c.id}.xlsx`), { force: true })
  }
  if (selected === INTERACTION_CASES) {
    for (const s of SEED_BUILDS) if (!seen.has(s.probe)) console.log(`${s.probe} (${s.id}): run --interaction --seeds (the seed build is the probe)`)
  }
  rmSync(WORK_DIR, { recursive: true, force: true })
}

async function interactionCli(args: string[]): Promise<number> {
  const ids = args.includes('--case') ? args[args.indexOf('--case') + 1]!.split(',') : undefined
  const tasks = args.includes('--task') ? args[args.indexOf('--task') + 1]!.split(',').map(Number) : undefined
  const unknown = ids?.filter(id => !INTERACTION_CASES.some(c => c.id === id)) ?? []
  if (unknown.length) { console.error(`unknown interaction case ${unknown.join(', ')}`); return 2 }
  const cases = INTERACTION_CASES.filter(c => (!ids || ids.includes(c.id)) && (!tasks || tasks.includes(c.task)))
  const seedFiles = new Set(cases.map(startOf))
  const needed = SEED_BUILDS.filter(s => seedFiles.has(s.file))
  if (args.includes('--dry-run')) return await interactionDryRun(cases, ids || tasks ? needed : SEED_BUILDS) ? 1 : 0
  if (!existsSync(ENGINE_SEED)) { console.error(`${relative(ROOT, ENGINE_SEED)} is missing: run --seed first`); return 2 }
  if (args.includes('--probes')) { await interactionProbes(ids || tasks ? cases : INTERACTION_CASES); return 0 }
  const fallback = !args.includes('--no-fallback')
  if (args.includes('--seeds')) {
    for (const s of SEED_BUILDS) await buildInteractionSeed(s, fallback)
    rmSync(WORK_DIR, { recursive: true, force: true })
    return 0
  }
  mkdirSync(INTERACTION_DIR, { recursive: true }); mkdirSync(WORK_DIR, { recursive: true })
  for (const s of needed) {
    if (existsSync(fixturePath(s.file))) continue
    try { await buildInteractionSeed(s, fallback) } catch (e) { console.log(`seed ${s.id} not built: ${message(e)}`) }
  }
  const entries = new Map(loadInteractionManifest().map(e => [e.case, e]))
  for (const c of cases) {
    entries.set(c.id, await buildInteractionCase(c, fallback))
    writeInteractionManifest([...entries.values()])   // after every case, so an interrupted batch keeps what it built
  }
  writeInteractionRecipes([...entries.values()])
  rmSync(WORK_DIR, { recursive: true, force: true })
  const built = cases.filter(c => entries.get(c.id)?.status === 'built')
  console.log(`${built.length}/${cases.length} built; manifest ${relative(ROOT, INTERACTION_MANIFEST)}`)
  for (const c of cases) { const e = entries.get(c.id); if (e?.status === 'pending') console.log(`  pending ${c.id}: ${e.error ?? ''}`) }
  return 0
}

// ---- CLI --------------------------------------------------------------------------------------------------------

if ((import.meta as { main?: boolean }).main && process.argv.includes('--interaction')) {
  process.exit(await interactionCli(process.argv.slice(2)))
} else if ((import.meta as { main?: boolean }).main) {
  const args = process.argv.slice(2)
  const ids = args.includes('--case') ? args[args.indexOf('--case') + 1]!.split(',') : undefined
  const unknown = ids?.filter(id => !ENGINE_CASES.some(c => c.id === id)) ?? []
  if (unknown.length) { console.error(`unknown case ${unknown.join(', ')}`); process.exit(2) }
  const cases = ENGINE_CASES.filter(c => !ids || ids.includes(c.id))
  if (args.includes('--dry-run')) process.exit(await dryRun(ids || !args.includes('--seed') ? cases : [], args.includes('--seed')) ? 1 : 0)
  if (args.includes('--seed')) { await buildSeed(); process.exit(0) }
  if (!existsSync(ENGINE_SEED)) { console.error(`${relative(ROOT, ENGINE_SEED)} is missing: run --seed first`); process.exit(2) }
  if (args.includes('--probes')) { await probes(); process.exit(0) }
  mkdirSync(WORK_DIR, { recursive: true })
  const entries = new Map(loadManifest().map(e => [e.case, e]))
  for (const c of cases) {
    entries.set(c.id, await buildCase(c, !args.includes('--no-fallback')))
    writeManifest([...entries.values()])   // after every case, so an interrupted batch keeps what it built
  }
  writeRecipes([...entries.values()])
  rmSync(WORK_DIR, { recursive: true, force: true })
  const built = cases.filter(c => entries.get(c.id)?.status === 'built').length
  console.log(`${built}/${cases.length} built; manifest ${relative(ROOT, MANIFEST)}`)
}
