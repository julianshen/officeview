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
  ENGINE_CASES, PROBES_TO_CONFIRM, SEED_CACHE_MARKERS, SEED_PIVOT_MARKERS, SOURCE_REF, engineSeedStandIn, seedDataLines, seedSourceLines, xmlDiffSummary,
  type EngineCase, type EngineRoute, type PivotParts, type Target,
} from './engine-cases'

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
const PIVOT_DIR = join(ROOT, 'tests/fixtures/xlsx-pivots')
const SEED_IN = join(PIVOT_DIR, 'recipes/pivot-seed.xlsx')
export const ENGINE_SEED = join(PIVOT_DIR, 'recipes/engine-seed.xlsx')
const OUT_DIR = join(PIVOT_DIR, 'engine')
const WORK_DIR = join(OUT_DIR, '.work')
const MANIFEST = join(OUT_DIR, 'manifest.json')
const RECIPES_DOC = join(ROOT, 'docs/validation/xlsx-pivots-recipes-stage2.md')
const PIVOT_RE = /^xl\/pivotTables\/pivotTable\d+\.xml$/, CACHE_RE = /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/

export interface ManifestEntry {
  case: string
  file: string
  features: string[]
  task: number
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
/** Opens `input`, runs `lines` then `refresh table`, saves as `saved` and closes. */
function session(input: string, saved: string, lines: (t: Target) => string[]): void {
  const name = openWorkbook(input)
  const t = target(name)
  try {
    runScript([...lines(t), `refresh table ${t.p}`], 180_000)
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

async function verify(file: string, markers: RegExp[], cacheMarkers: RegExp[] = [], absent: RegExp[] = []): Promise<string> {
  const parts = await readParts(file)
  for (const m of markers) if (!m.test(parts.pivot)) throw new Error(`saved pivot lacks ${m}`)
  for (const m of cacheMarkers) if (!m.test(parts.cache)) throw new Error(`saved cache lacks ${m}`)
  for (const m of absent) if (m.test(parts.pivot)) throw new Error(`saved pivot has ${m}`)
  return parts.pivotPart
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
    try { await verify(ENGINE_SEED, SEED_PIVOT_MARKERS, SEED_CACHE_MARKERS); console.log(`built ${relative(ROOT, ENGINE_SEED)} (applescript)`); return }
    catch (e) { console.log(`seed: ${message(e)}; taking the xml route`); copyFileSync(ENGINE_SEED, staged) }
  }
  await writeEdited(staged, edited, p => ({ ...p, cache: p.cache.replace(/(<worksheetSource [^>]*ref=")[^"]*"/, `$1${SOURCE_REF}"`) }))
  session(edited, ENGINE_SEED, () => [])
  await verify(ENGINE_SEED, SEED_PIVOT_MARKERS, SEED_CACHE_MARKERS)
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
  const entry = (route: EngineRoute, status: ManifestEntry['status'], pivot = 'xl/pivotTables/pivotTable1.xml', error?: string): ManifestEntry =>
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

// ---- CLI --------------------------------------------------------------------------------------------------------

if ((import.meta as { main?: boolean }).main) {
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
