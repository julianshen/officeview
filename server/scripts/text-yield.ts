/** Stage 3a Task 9 — visible-text yield over the corpus (docs/validation/streaming-s3a.md).
 *
 * For every document in `corpus/` and `corpus-s0/` that the render pool opens
 * (`contentPolicy: 'visible-only'`, the protected default), run `prepare` on every
 * unit through `Pool.text` — the production path: admission, the worker's grant
 * check, the recording and visual pass, the R1 join — and tally per format the
 * spans served and withheld, the withheld spans by cause (the join's diagnostic
 * `causes` tally, counts only), and units refused for text (grant, oversize,
 * timeouts, admission) separately. Units are prepared one at a time with the
 * default caps, so the yield is not distorted by load-made refusals.
 *
 *   bun --preload ./scripts/stream-s0/napi-preload.ts server/scripts/text-yield.ts \
 *     [--files REGEX] [--rows rows.json]
 *
 * Run from the repo root. `--rows` writes one row per unit (document, unit,
 * served, withheld, causes or refusal) for picking example documents.
 * Exit 1 if any format serves nothing. */
import { readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { AdmissionRefused, Pool, RenderFailed, RenderTimeout } from '../src/pool'
import type { WithheldCauses } from '../src/text/join'
import type { TextReply } from '../src/text/served'

type Prepared = Extract<TextReply, { op: 'prepare' }>
interface Row { doc: string; format: string; unit: number; ms: number; served?: number; withheld?: number; causes?: WithheldCauses; refused?: string }
interface Tally { docs: number; notOpened: number; units: number; prepared: number; refused: Record<string, number>; served: number; withheld: number; causes: Record<string, number> }

const DIRS = ['corpus', 'corpus-s0']
const FORMATS = ['docx', 'xlsx', 'pptx'] as const

/** Cause keys as printed: `no candidate run`, `occluded:<reason>`, `character mismatch`, `whitespace`, `bad scale`. */
export function causeEntries(c: WithheldCauses): Array<[string, number]> {
  const out: Array<[string, number]> = [['no candidate run', c.noCandidate]]
  for (const [reason, n] of Object.entries(c.occluded)) out.push([`occluded:${reason}`, n ?? 0])
  out.push(['character mismatch', c.mismatch], ['whitespace', c.whitespace], ['bad scale', c.badScale])
  return out.filter(([, n]) => n > 0)
}

function refusalOf(e: unknown): string {
  if (e instanceof AdmissionRefused) return `admission:${e.reason}`
  if (e instanceof RenderTimeout) return 'timeout'
  if (e instanceof RenderFailed) return e.refusal ? `refused:${e.refusal}` : `failed:${e.message.slice(0, 60)}`
  return `error:${String(e).slice(0, 60)}`
}

const pct = (a: number, b: number) => (b ? `${((100 * a) / b).toFixed(1)} %` : 'n/a')

async function main(args: string[]): Promise<number> {
  const opt = (name: string) => (args.includes(name) ? args[args.indexOf(name) + 1] : undefined)
  const filter = opt('--files') ? new RegExp(opt('--files')!) : undefined
  const rowsOut = opt('--rows')
  const root = resolve(import.meta.dir, '../..')
  const files = DIRS.flatMap((dir) => {
    let names: string[] = []
    try { names = readdirSync(resolve(root, dir)) } catch { /* absent directory: none */ }
    return names.filter((n) => /\.(docx|xlsx|pptx)$/i.test(n) && !n.startsWith('~$') && (!filter || filter.test(n))).map((n) => { const path = resolve(root, dir, n); return { dir, name: n, path, size: statSync(path).size } })
  }).sort((a, b) => b.size - a.size)   // largest first: a slow unit shows early (sizes read once, by stat)
  const log = (s: string) => process.stderr.write(`${new Date().toISOString().slice(11, 19)} ${s}\n`)
  log(`${files.length} documents (${DIRS.map((d) => `${d}: ${files.filter((f) => f.dir === d).length}`).join(', ')})`)

  const pool = new Pool({ workers: 1 })
  const tally = Object.fromEntries(FORMATS.map((f) => [f, { docs: 0, notOpened: 0, units: 0, prepared: 0, refused: {}, served: 0, withheld: 0, causes: {} } as Tally])) as Record<string, Tally>
  const rows: Row[] = []
  const notOpened: Array<{ doc: string; why: string }> = []
  const t0 = performance.now()
  try {
    for (const [i, f] of files.entries()) {
      const format = f.name.split('.').pop()!.toLowerCase(), t = tally[format]!
      t.docs++
      const docRef = `yield-${i}`
      let units: number
      try { units = (await pool.open(docRef, new Uint8Array(readFileSync(f.path)), 'visible-only')).units.length } catch (e) {
        t.notOpened++; notOpened.push({ doc: f.name, why: refusalOf(e) }); log(`${f.name}: not opened (${refusalOf(e)})`); continue
      }
      const d0 = performance.now()
      let served = 0, withheld = 0
      for (let u = 0; u < units; u++) {
        const u0 = performance.now()
        t.units++
        try {
          const r = await pool.text(`yield-${i}`, { docRef, contentPolicy: 'visible-only', op: 'prepare', unit: u }) as Prepared
          t.prepared++; t.served += r.served; t.withheld += r.withheld; served += r.served; withheld += r.withheld
          for (const [k, n] of causeEntries(r.causes)) t.causes[k] = (t.causes[k] ?? 0) + n
          rows.push({ doc: f.name, format, unit: u, ms: Math.round(performance.now() - u0), served: r.served, withheld: r.withheld, causes: r.causes })
        } catch (e) {
          const why = refusalOf(e)
          t.refused[why] = (t.refused[why] ?? 0) + 1
          rows.push({ doc: f.name, format, unit: u, ms: Math.round(performance.now() - u0), refused: why })
        }
      }
      pool.dropParsed(docRef, 'visible-only')
      log(`${f.name}: ${units} units, served ${served}, withheld ${withheld} (${Math.round(performance.now() - d0)} ms)`)
    }
  } finally {
    pool.dispose()
  }

  const lines = [`text yield — visible-only, prepare per unit, ${((performance.now() - t0) / 1000).toFixed(1)} s`, '']
  lines.push('| format | docs (not opened) | units (prepared / refused) | spans | served | withheld | served excl. withheld blank spans | top withheld causes (share of spans) |', '|---|---|---|---|---|---|---|---|')
  let empty = false
  for (const f of FORMATS) {
    const t = tally[f]!, spans = t.served + t.withheld
    if (t.docs && !t.served) empty = true
    const refused = Object.values(t.refused).reduce((a, b) => a + b, 0)
    const causes = Object.entries(t.causes).sort((a, b) => b[1] - a[1]).map(([k, n]) => `${k} ${n} (${pct(n, spans)})`).join('; ') || '—'
    lines.push(`| ${f} | ${t.docs} (${t.notOpened}) | ${t.units} (${t.prepared} / ${refused}) | ${spans} | ${t.served} (${pct(t.served, spans)}) | ${t.withheld} | ${pct(t.served, spans - (t.causes.whitespace ?? 0))} | ${causes} |`)
  }
  lines.push('')
  for (const f of FORMATS) {
    const r = Object.entries(tally[f]!.refused)
    if (r.length) lines.push(`refused ${f}: ${r.map(([k, n]) => `${k} ${n}`).join(', ')}`)
  }
  for (const n of notOpened) lines.push(`not opened: ${n.doc} (${n.why})`)
  process.stdout.write(`${lines.join('\n')}\n`)
  if (rowsOut) writeFileSync(rowsOut, JSON.stringify(rows, null, 1))
  return empty ? 1 : 0
}

if (import.meta.main) process.exit(await main(process.argv.slice(2)))
