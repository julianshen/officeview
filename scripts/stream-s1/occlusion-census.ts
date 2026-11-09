/**
 * Stage 1d Task 5 — §7.7 occlusion-evidence census over the Stage-0 corpus.
 *
 *   bun --preload ./scripts/stream-s0/napi-preload.ts scripts/stream-s1/occlusion-census.ts
 *
 * Every unit is rendered through the server's DocumentHost (visible-only,
 * final-raster, scale 1, mark off) and its evidence is summed per kind:
 * text runs, servable, withheld, and the withholding reasons. Evidence only —
 * nothing is culled or shipped in Stage 1.
 */
import { readFileSync } from 'node:fs'
import { DocumentHost, RenderRefusal } from '../../server/src/render'
import { listCorpus } from '../stream-s0/corpus'

type Row = { files: number; units: number; refused: number; runs: number; servable: number; withheld: number; reasons: Record<string, number> }
const rows: Record<string, Row> = {}
const host = new DocumentHost()
for (const file of listCorpus()) {
  const row = rows[file.kind] ??= { files: 0, units: 0, refused: 0, runs: 0, servable: 0, withheld: 0, reasons: {} }
  let opened
  try { opened = await host.open(file.name, new Uint8Array(readFileSync(file.path)), 'visible-only') } catch { row.refused++; continue }
  row.files++
  for (let i = 0; i < opened.units.length; i++) {
    try {
      const e = host.render(opened.docKey, { unitIndex: i, scale: 1, projection: 'final-raster', mark: { kind: 'off' } }).evidence!
      row.units++; row.runs += e.runs; row.servable += e.servable; row.withheld += e.withheld
      for (const [k, v] of Object.entries(e.reasons)) row.reasons[k] = (row.reasons[k] ?? 0) + (v ?? 0)
    } catch (err) { if (err instanceof RenderRefusal) row.refused++; else throw err }
  }
  host.drop(opened.docKey)
}
console.log('| kind | files | units | text runs | servable | withheld | reasons |')
console.log('|---|---|---|---|---|---|---|')
for (const [kind, r] of Object.entries(rows)) {
  const reasons = Object.entries(r.reasons).sort((a, b) => b[1] - a[1]).map(([k, v]) => `${k} ${v}`).join(', ') || '—'
  console.log(`| ${kind} | ${r.files} | ${r.units}${r.refused ? ` (+${r.refused} refused)` : ''} | ${r.runs} | ${r.servable} (${(100 * r.servable / Math.max(1, r.runs)).toFixed(1)} %) | ${r.withheld} | ${reasons} |`)
}
