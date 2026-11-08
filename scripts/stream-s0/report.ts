/**
 * Task 11: renders every docs/validation/streaming-s0/*.json gate result into
 * docs/validation/streaming.md (spec §17 row 0: "all gates green with
 * published numbers … any gate failing reports to the owner").
 *
 *   bun run s0:report
 *
 * A missing result is listed as NOT RUN, never omitted.
 */
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

type Json = Record<string, any>

export interface GateRow { id: string; title: string; status: string; headline: string; target: string }

export const GATES: ReadonlyArray<{ id: string; file: string; title: string; target: string; headline: (d: Json) => string }> = [
  { id: 'T1', file: 't1', title: 'Backend confirm (@napi-rs/canvas)', target: 'no napi-only failures; 9/9 construction sites ok',
    headline: (d) => `${d.renderSuccess?.okNapi}/${d.renderSuccess?.okNode} renders ok on napi, napi-only failures ${len(d.renderSuccess?.napiOnlyFailures)}, sites ok ${(d.constructionSites ?? []).filter((s: Json) => s.napi === 'ok').length}/${(d.constructionSites ?? []).length}` },
  { id: 'T2', file: 't2', title: 'Hidden-content matrix (Stage-0 subset)', target: 'no hidden ink drawn',
    headline: (d) => `${d.totals?.leaks} of ${d.totals?.fixtures} hidden constructs drawn by today's renderer (${d.totals?.invalid} invalid)` },
  { id: 'T3', file: 't3', title: 'Recorder census', target: 'notInSubset empty',
    headline: (d) => `${d.unitsSeen} units, ${Object.keys(d.methods ?? {}).length} methods, notInSubset ${len(d.notInSubset)}` },
  { id: 'T4', file: 't4', title: 'Font probes', target: 'reported',
    headline: (d) => `${pct(d.totals?.rasterFallbackRate)} of units would raster (${d.totals?.unitsThatWouldRaster}/${d.totals?.unitsTotal}); ${d.totals?.unverifiable}/${d.totals?.faces} faces unverifiable; with the pinned open pack ${pct(d.totals?.rasterFallbackRateWithPack)} (${d.totals?.unitsThatWouldRasterWithPack}/${d.totals?.unitsTotal})` },
  { id: 'T5', file: 't5', title: 'Show-plan client prototype', target: 'package-free client; static slides ≤ 0.1 % diff',
    headline: (d) => `eligibility ${pct(d.eligibilityRate)} (${d.proofSubset}/${d.steps}); static ${d.staticSlides} slides, ${d.staticFailures} failures` },
  { id: 'T7', file: 't7', title: 'Layer-1 mark codec + ΔE00', target: 'page ΔE00 max ≤ 3',
    headline: (d) => {
      const row = (d.sweep ?? []).find((r: Json) => JSON.stringify(r.params) === JSON.stringify(d.chosen))
      return `chosen ${JSON.stringify(d.chosen)}; page ΔE00 max ${num(row?.pageMax)}; q70→½ pilot ${pct(row?.q70HalfRate)}`
    } },
  { id: 'T8', file: 't8', title: 'XLSX tile timing', target: 'tile p95 ≤ 120 ms',
    headline: (d) => `best-of-3 p95 ${num(d.bestOf3P95)} ms (${d.sheet?.rows}×${d.sheet?.cols}, ${d.tilePx} px tiles)` },
  { id: 'T9', file: 't9', title: 'Detector calibration + survival', target: '≤ 1e-6 false attribution; §13.5 survival table',
    headline: (d) => `τ_abs ${num(d.tau?.abs)}, ${d.fullScoredDecisions} null decisions, false ${d.falseCount}; RED rows: ${(d.survival ?? []).filter((r: Json) => r.status === 'RED').map((r: Json) => `${r.attack} ${pct(r.rate)}`).join(', ') || 'none'}` },
  { id: 'T11', file: 't11', title: 'Eager gates on the current build', target: 'transitive gzip ≤ 38,720 B',
    headline: (d) => `closure ${d.transitiveGzipBytes} B over ${len(d.perChunk)} chunks (index.js alone ${d.indexOnlyGzipBytes} B)` },
  { id: 'T12', file: 't12', title: 'Synthetic admission probe', target: 'no §14 cap exceeded',
    headline: (d) => `${d.requests?.total} requests, peak reserved ${mib(d.peak?.reservedBytes)}, refused ${d.requests?.refused}, p95 ${num(d.latencyMs?.p95, 0)} ms, RSS ${mib(d.peak?.rssBytes)}` },
]

function len(v: unknown): number { return Array.isArray(v) ? v.length : typeof v === 'number' ? v : 0 }
function pct(v: unknown): string { return typeof v === 'number' ? `${(v * 100).toFixed(1)} %` : '—' }
function num(v: unknown, digits = 2): string { return typeof v === 'number' ? v.toFixed(digits) : '—' }
function mib(v: unknown): string { return typeof v === 'number' ? `${(v / 2 ** 20).toFixed(1)} MiB` : '—' }

/** Non-gate result files the report also reads. */
export const SUPPLEMENTS = ['image-heavy', 't9-image-heavy'] as const

/** Owner rulings on RED gates (spec §2 item 14, draft 12). */
export const RULINGS: Record<string, string> = {
  T2: 'expected baseline — the six drawn constructs are the confirmed §6.4 server strip list, the strip fails closed on unrecognised hidden flags, and zero hidden ink on the stream path gates Stage 1',
  T11: 'the 380 KB chunk is the shared parser/renderer engine `<OfficeDoc>` imports eagerly (removing the worker re-export was measured and does not change it); the 38,720 B transitive ceiling applies to the stream client entry `officeview/core`, and the main entry keeps the index.js-alone budget plus a no-growth gate on its closure (§8.1, §2 item 14b)',
}

export function loadResults(dir: string): Map<string, Json | undefined> {
  const out = new Map<string, Json | undefined>()
  for (const g of GATES) {
    const p = join(dir, `${g.file}.json`)
    out.set(g.id, existsSync(p) ? JSON.parse(readFileSync(p, 'utf8')) as Json : undefined)
  }
  for (const extra of SUPPLEMENTS) {
    const p = join(dir, `${extra}.json`)
    out.set(extra, existsSync(p) ? JSON.parse(readFileSync(p, 'utf8')) as Json : undefined)
  }
  return out
}

export function gateRows(results: Map<string, Json | undefined>): GateRow[] {
  return GATES.map((g) => {
    const d = results.get(g.id)
    if (!d) return { id: g.id, title: g.title, status: 'NOT RUN', headline: '—', target: g.target }
    let headline: string
    try { headline = g.headline(d) } catch { headline = '(unreadable result)' }
    return { id: g.id, title: g.title, status: String(d.status), headline, target: g.target }
  })
}

const SCOPE_CALLS = `1. **T11** — Stage 0 runs the §8.1 gate scripts against the current build and publishes the numbers plus a module-map baseline hash. Building \`officeview/core\` and its purity assert are Stage 1.
2. **T12** — no server exists at Stage 0, so T12 is a synthetic admission probe: an in-process model of the pool's reservation/queue rules driven by 1,000 concurrent requests (plus 5 rogue sessions × 100) with real \`@napi-rs/canvas\` paints in worker threads.
3. **T2** — Stage-0 subset = hidden-content fixture matrix + paint-leak check + model-exposure inventory. The Playwright WS harness and "local loads byte-for-byte unchanged" are Stage 1.
4. **T6** has no Stage-0 subset (wire is Stage 1). **T8** Stage-0 subset = tile p95 under \`@napi-rs/canvas\`.
5. **T1** excludes "fractional zoom proves 1:1 replay" (needs the recorder, Stage 1).
6. **T7/T9** cover Layer 1 only; Layer 2 is command projection, Stage 2+.`

const DECISIONS = `- **Keyed non-linear mark message** (spec §13.2, commit 391616b6): a Stage-0 review showed a K = 5 median of colluders framed an innocent token under the linear BCH/CRC map (innocent Z 9.19–9.40). Now \`msg_40 = HMAC(kMsg, keyGen | codeword_40)[:40]\`, and ≥ 2 candidates over τ ⇒ \`collusion-suspected\` (names no one). Measured innocent Z after the fix: −0.06.
- **"25 % crop" = 25 % removed from each axis** (≈ 44 % of the area) (spec §13.5, commit 8420b855).
- **T7 ΔE00 choice gates on page max**, not p95 — marks cover < 1 % of a real page, so p95 is ~0 (plan, commit 767e26ec).
- **Stage-0 rulings (spec draft 12, §2 item 14):** T2 strip list confirmed + fail-closed rule (§6.4); T11 ceiling applies to the \`officeview/core\` stream client, main entry no-growth (§8.1); a pinned open font pack of metric-compatible faces (§7.4 rule 6); DOCX top-up and hand-certified image-heavy units with a T9 supplement run.`

export function renderReport(results: Map<string, Json | undefined>): string {
  const rows = gateRows(results)
  const red = rows.filter((r) => r.status === 'RED' || r.status === 'NOT RUN')
  const t1 = results.get('T1'); const t9 = results.get('T9'); const t4 = results.get('T4'); const t2 = results.get('T2'); const t3 = results.get('T3'); const t5 = results.get('T5'); const t11 = results.get('T11')
  const lines: string[] = []
  lines.push('# Server streaming — Stage 0 feasibility results', '')
  lines.push('Generated by `bun run s0:report` from `docs/validation/streaming-s0/*.json`. Spec: `docs/superpowers/specs/2026-11-07-server-streaming-design.md` §16–17. Plan: `docs/superpowers/plans/2026-11-08-streaming-stage0.md`.', '')
  lines.push('## Gate summary', '', '| Gate | Status | Headline | Target |', '|---|---|---|---|')
  for (const r of rows) lines.push(`| ${r.id} ${r.title} | **${r.status}** | ${r.headline} | ${r.target} |`)
  lines.push('', '## RED gates — owner decision needed', '')
  if (red.length === 0) lines.push('None.')
  for (const r of red) lines.push(`- **${r.id} ${r.title}: ${r.status}** — ${r.headline} (target: ${r.target}).${RULINGS[r.id] ? ` **Owner ruling:** ${RULINGS[r.id]}.` : ''}`)
  lines.push('', '## Owner decisions taken during Stage 0', '', DECISIONS)
  lines.push('', '## Stage-0 scope calls', '', SCOPE_CALLS)
  if (t1) {
    lines.push('', '## T1 cross-backend tolerance (node-canvas vs @napi-rs/canvas)', '', '| kind | p95 fraction of pixels differing | max channel Δ |', '|---|---|---|')
    for (const [k, v] of Object.entries(t1.tolerance?.byKind ?? {}) as Array<[string, Json]>) lines.push(`| ${k} | ${pct(v.p95FracDiffering)} | ${v.maxChannelDelta} |`)
  }
  if (t3) lines.push('', `## T3 census`, '', `notInSubset: ${len(t3.notInSubset) === 0 ? 'empty — the §7.3/§7.5 recorded subset covers every op the corpus paints' : JSON.stringify(t3.notInSubset)}. Files that failed to load: ${(t3.unitsFailed ?? []).map((f: Json) => f.file).join(', ') || 'none'}.`)
  if (t4) {
    lines.push('', '## T4 fonts — families that force raster fallback', '')
    lines.push(`Without a pack (counts are files): ${(t4.familiesCausingRaster ?? []).map((f: Json) => `${f.family} (${f.count})`).join(', ')}.`)
    if (t4.familiesCausingRasterWithPack) {
      lines.push('', `With the pinned open pack (${(t4.pinnedPack ?? []).map((p: Json) => `${p.family} → ${p.file}`).join(', ')}): raster rate ${pct(t4.totals?.rasterFallbackRateWithPack)}; remaining causes ${(t4.familiesCausingRasterWithPack ?? []).map((f: Json) => `${f.family} (${f.count})`).join(', ')}. No metric-compatible open face exists for those.`)
    }
  }
  if (t2) {
    lines.push('', '## T2 hidden content drawn by today\'s renderer (the §6.4 strip list)', '', '| fixture | status | model paths |', '|---|---|---|')
    for (const f of t2.perFixture ?? []) lines.push(`| ${f.id} | ${f.status} | ${(f.modelExposure ?? []).slice(0, 3).join('<br>')} |`)
  }
  if (t5) lines.push('', '## T5 presentation', '', `Appearance eligibility ${pct(t5.eligibilityRate)} (${t5.proofSubset}/${t5.steps} steps; conservative — no swept-area proof yet). Coverage gaps: ${JSON.stringify(t5.coverageGaps ?? [])}.`)
  if (t9) {
    lines.push('', '## T9 survival (image-only units, chosen parameters)', '', '| attack | correct | rate | target | status |', '|---|---|---|---|---|')
    for (const r of t9.survival ?? []) lines.push(`| ${r.attack} | ${r.correct}/${r.total} | ${pct(r.rate)} | ${r.target == null ? '—' : pct(r.target)} | ${r.status} |`)
    lines.push('', `Calibration: τ_abs ${num(t9.tau?.abs)} / τ_margin ${num(t9.tau?.margin)} at M = ${t9.candidateCount}; τ at M = 2²⁴ ≈ ${num(t9.tau24)}. ${t9.fullScoredDecisions} full null decisions, ${t9.falseCount} false attributions (one-sided 95 % upper bound ${t9.zeroFalse95Upper}; 1e-6 rests on the fitted tail and union bound, not on observation).`)
    lines.push('', '## T9 collusion', '', '| attack | colluder named | collusion-suspected | innocent named | none |', '|---|---|---|---|---|')
    for (const c of t9.collusion ?? []) lines.push(`| ${c.attack} | ${c.colluder} | ${c.suspected} | ${c.innocent} | ${c.none} |`)
  }
  const ih = results.get('image-heavy'); const t9ih = results.get('t9-image-heavy')
  if (ih) {
    lines.push('', '## Corpus top-up and image-heavy certification', '')
    lines.push(`Gates T1–T9 above ran on the original 46-file corpus. The top-up adds 6 POI DOCX files and 5 POI decks (\`corpus-s0/\` via \`scripts/stream-s0/fetch-extra.ts\`, outside the shared corpus). Image-heavy = image pixel share ≥ ${pct(ih.threshold?.imageShare)} (paint with vs without \`drawImage\`) **and** texture ≥ ${pct(ih.threshold?.textureShare)} (8×8 blocks with luma σ ≥ ${ih.threshold?.blockLumaStd}). ${len(ih.certified)} of ${len(ih.units)} units certified (minimum 5) — ${ih.status}. The 11 LibreOffice slides that are 70–100 % image all fail the texture test (flat backgrounds).`)
  }
  if (t9ih) {
    lines.push('', `## T9 supplement — image-heavy and DOCX top-up units (τ from the full run, not recalibrated)`, '', `Units: ${(t9ih.units?.imageHeavy ?? []).length} image-heavy (${(t9ih.units?.imageHeavy ?? []).join(', ')}) + ${(t9ih.units?.docxTopup ?? []).length} DOCX (${(t9ih.units?.docxTopup ?? []).join(', ')}), bringing T9's DOCX units to 15. Status ${t9ih.status}.`, '', '| attack | image-heavy | DOCX top-up | target | status |', '|---|---|---|---|---|')
    const byClass = (cls: string, attack: string) => (t9ih.survivalByClass ?? []).find((c: Json) => c.class === cls)?.rows?.find((r: Json) => r.attack === attack)
    for (const r of t9ih.survival ?? []) {
      const a = byClass('image-heavy', r.attack); const b = byClass('docx-topup', r.attack)
      lines.push(`| ${r.attack} | ${a ? `${a.correct}/${a.total} (${pct(a.rate)})` : '—'} | ${b ? `${b.correct}/${b.total} (${pct(b.rate)})` : '—'} | ${r.target == null ? '—' : pct(r.target)} | ${r.status} |`)
    }
    lines.push('', '| collusion | colluder named | collusion-suspected | innocent named | none |', '|---|---|---|---|---|')
    for (const c of t9ih.collusion ?? []) lines.push(`| ${c.attack} | ${c.colluder} | ${c.suspected} | ${c.innocent} | ${c.none} |`)
  }
  if (t11) lines.push('', '## T11 note', '', `\`dist/index.js\` statically imports ${(t11.perChunk ?? []).map((c: Json) => `\`${c.file}\` (${c.gzipBytes} B)`).join(', ')}. Today's \`eager-budget.ts\` measures index.js alone (${t11.indexOnlyGzipBytes} B ≤ 38,720). §8.1 gate 1 as written (the whole static closure) is ${t11.transitiveGzipBytes} B — the gap pre-dates this branch.`)
  lines.push('')
  return lines.join('\n')
}

if (import.meta.main) {
  const root = join(fileURLToPath(new URL('.', import.meta.url)), '..', '..')
  const md = renderReport(loadResults(join(root, 'docs', 'validation', 'streaming-s0')))
  writeFileSync(join(root, 'docs', 'validation', 'streaming.md'), md)
  console.log(md.split('\n').slice(0, 20).join('\n'))
}
