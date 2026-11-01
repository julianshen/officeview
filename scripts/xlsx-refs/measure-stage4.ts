/**
 * Stage 4 measurements read off the Excel-saved oracle files (pivots plan Task 34, step 4), printed as markdown for
 * docs/validation/xlsx-pivots.md ("Stage 4 interaction oracle and measurements"). Read-only; run after
 * `build-engine.ts --interaction --task 35,36,37`.
 *
 *   bun scripts/xlsx-refs/measure-stage4.ts [--case id[,id…]] [--cells]
 *
 * Per case: the pivot's location and axes, `dataOnRows`/`dataPosition`, every data field, the pivot fields that
 * differ from the start file, the cache fields that are grouped or new, the filters, and (with --cells) the Pivot
 * sheet's cells inside the location.
 */
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { startOf, type InteractionCase } from './interaction-cases'
import type { InteractionManifestEntry } from './build-engine'
import { CACHE_RE, PIVOT_DIR, PIVOT_RE, fixturePath, loadZip, partNames, partText, sharedStrings, sheetCells, sheetXml } from './pivot-files'

const manifestFile = join(PIVOT_DIR, 'interaction/manifest.json')
const all = existsSync(manifestFile) ? JSON.parse(readFileSync(manifestFile, 'utf8')) as InteractionManifestEntry[] : []

interface Parts { pivot: string; cache: string; cells: Array<{ ref: string; value?: string }> }
async function parts(file: string): Promise<Parts> {
  const zip = await loadZip(file)
  const pivot = (await partText(zip, partNames(zip, PIVOT_RE)[0]))!, cache = (await partText(zip, partNames(zip, CACHE_RE)[0]))!
  const xml = await sheetXml(zip, 'Pivot')
  return { pivot, cache, cells: xml ? sheetCells(xml, await sharedStrings(zip)) : [] }
}
const one = (xml: string, re: RegExp) => re.exec(xml)?.[0] ?? '—'
const attr = (tag: string, name: string) => new RegExp(`\\s${name}="([^"]*)"`).exec(tag)?.[1]
/** The `<pivotField>` elements in order (self-closing or with children). */
const pivotFields = (pivot: string) => [...(/<pivotFields[^>]*>([\s\S]*?)<\/pivotFields>/.exec(pivot)?.[1] ?? '').matchAll(/<pivotField\b[^>]*?(?:\/>|>[\s\S]*?<\/pivotField>)/g)].map(m => m[0])
const cacheFields = (cache: string) => [...cache.matchAll(/<cacheField\b[^>]*?(?:\/>|>[\s\S]*?<\/cacheField>)/g)].map(m => m[0])
const clip = (s: string, n = 400) => s.length > n ? `${s.slice(0, n)}… (${s.length} chars)` : s
const openTag = (el: string) => /^<[^>]*>/.exec(el)?.[0] ?? el

function inLocation(ref: string, location: string): boolean {
  const [a, b] = location.split(':')
  const rc = (r: string) => { const m = /^([A-Z]+)(\d+)$/.exec(r)!; return [Number(m[2]), [...m[1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)] as const }
  const [r, c] = rc(ref), [r0, c0] = rc(a!), [r1, c1] = rc(b ?? a!)
  return r >= r0 && r <= r1 && c >= c0 && c <= c1
}

async function report(e: InteractionManifestEntry, withCells: boolean): Promise<string> {
  const start = await parts(fixturePath(startOf(e as Pick<InteractionCase, 'seed' | 'base'>)))
  const saved = await parts(fixturePath(`interaction/${e.case}`))
  const out = [`#### ${e.case} (Task ${e.task}; start ${startOf(e as Pick<InteractionCase, 'seed' | 'base'>)}; route ${e.route})`, '', `- measures: ${e.measures}`]
  const location = /<location [^>]*ref="([^"]+)"/.exec(saved.pivot)?.[1] ?? ''
  out.push(`- location: \`${one(start.pivot, /<location [^>]*\/>/)}\` → \`${one(saved.pivot, /<location [^>]*\/>/)}\``)
  const def = openTag(one(saved.pivot, /<pivotTableDefinition [^>]*>/))
  out.push(`- dataOnRows=${attr(def, 'dataOnRows') ?? '—'} dataPosition=${attr(def, 'dataPosition') ?? '—'} rowGrandTotals=${attr(def, 'rowGrandTotals') ?? '—'} colGrandTotals=${attr(def, 'colGrandTotals') ?? '—'}`)
  for (const axis of ['rowFields', 'colFields', 'pageFields', 'dataFields', 'filters']) {
    const before = one(start.pivot, new RegExp(`<${axis}\\b[\\s\\S]*?</${axis}>`)), after = one(saved.pivot, new RegExp(`<${axis}\\b[\\s\\S]*?</${axis}>`))
    out.push(`- ${axis}: ${before === after ? `unchanged \`${clip(after, 300)}\`` : `\`${clip(before, 300)}\` → \`${clip(after, 500)}\``}`)
  }
  const [pf0, pf1] = [pivotFields(start.pivot), pivotFields(saved.pivot)]
  for (let k = 0; k < Math.max(pf0.length, pf1.length); k++) {
    if (pf0[k] === pf1[k]) continue
    out.push(`- pivotField ${k}: \`${clip(pf0[k] ?? '(none)', 300)}\` → \`${clip(pf1[k] ?? '(none)', 600)}\``)
  }
  const [cf0, cf1] = [cacheFields(start.cache), cacheFields(saved.cache)]
  for (let k = 0; k < Math.max(cf0.length, cf1.length); k++) {
    const a = cf0[k], b = cf1[k]
    if (a === b) continue
    const grouped = (b ?? '').includes('<fieldGroup') || (a ?? '').includes('<fieldGroup') || !a || !b
    const shared = (s?: string) => s ? openTag(one(s, /<sharedItems\b[^>]*\/?>/)) : '(none)'
    if (grouped) out.push(`- cacheField ${k}: \`${clip(a ? openTag(a) : '(none)', 200)}\` ${a?.includes('<fieldGroup') ? `\`${clip(one(a, /<fieldGroup[\s\S]*?(?:<\/fieldGroup>|\/>)/), 300)}\`` : ''} → \`${clip(b ? openTag(b) : '(none)', 200)}\` ${b?.includes('<fieldGroup') ? `\`${clip(one(b, /<fieldGroup[\s\S]*?(?:<\/fieldGroup>|\/>)/), 1200)}\`` : ''}`)
    else out.push(`- cacheField ${k} sharedItems: \`${clip(shared(a), 300)}\` → \`${clip(shared(b), 300)}\` (${(a?.match(/<[snmdbe] /g) ?? []).length} → ${(b?.match(/<[snmdbe][ /]/g) ?? []).length} items) \`${clip(one(b ?? '', /<sharedItems\b[^>]*\/>|<sharedItems\b[^>]*[^/]>[\s\S]*?<\/sharedItems>/), 900)}\``)
  }
  if (withCells && location) {
    const cells = saved.cells.filter(c => inLocation(c.ref, location) && c.value !== undefined)
    out.push(`- cells (${cells.length} in ${location}): ${clip(cells.map(c => `${c.ref}=${JSON.stringify(c.value)}`).join(' '), 2500)}`)
  }
  out.push('')
  return out.join('\n')
}

if ((import.meta as { main?: boolean }).main) {
  const args = process.argv.slice(2)
  const ids = args.includes('--case') ? args[args.indexOf('--case') + 1]!.split(',') : undefined
  const entries = all.filter(e => e.task >= 35 && e.status === 'built' && (!ids || ids.includes(e.case)))
  for (const e of entries) console.log(await report(e, args.includes('--cells')))
}
