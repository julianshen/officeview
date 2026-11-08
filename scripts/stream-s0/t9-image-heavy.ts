/**
 * T9 supplement (owner ruling 2026-11-08, spec §2 item 14d): survival and
 * collusion on the certified image-heavy units plus the DOCX top-up units,
 * with the chosen parameters and the τ calibrated by the full T9 run
 * (`t9.json`) — no recalibration, no parameter sweep re-selection.
 *
 *   bun run s0:napi scripts/stream-s0/t9-image-heavy.ts
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { listCorpus, loadUnits, renderUnit } from './corpus'
import { resolveBackend, writeResult } from './results'
import { runBattery, type Page } from './t7-t9-battery'
import type { MarkParams } from './mark/types'

const DIR = join(process.cwd(), 'docs', 'validation', 'streaming-s0')
/** One page of each top-up DOCX — brings T9's DOCX units from 12 to 15. */
export const DOCX_TOPUP = ['poi-heading123.docx#0', 'poi-bib-chernigovka.netdo.ru_download_docs_17459.docx#0', 'poi-drawing.docx#0']
const PER_FILE = 2

async function main(): Promise<number> {
  if (resolveBackend() !== '@napi-rs/canvas') {
    console.error('[t9-ih] refused: run with `bun run s0:napi scripts/stream-s0/t9-image-heavy.ts`')
    return 2
  }
  const t9 = JSON.parse(readFileSync(join(DIR, 't9.json'), 'utf8')) as { tau: { abs: number; margin: number } }
  const t7 = JSON.parse(readFileSync(join(DIR, 't7.json'), 'utf8')) as { chosen: MarkParams }
  const ih = JSON.parse(readFileSync(join(DIR, 'image-heavy.json'), 'utf8')) as { certified: Array<{ id: string; textureShare: number }> }
  // Top PER_FILE units per deck by texture, so one deck cannot dominate.
  const byFile = new Map<string, Array<{ id: string; textureShare: number }>>()
  for (const u of ih.certified) {
    const file = u.id.split('#')[0]!
    byFile.set(file, [...(byFile.get(file) ?? []), u])
  }
  const imageHeavy = [...byFile.values()].flatMap((list) => list.sort((a, b) => b.textureShare - a.textureShare).slice(0, PER_FILE).map((u) => u.id))
  const wanted = new Set([...imageHeavy, ...DOCX_TOPUP])
  const pages: Page[] = []
  for (const file of listCorpus()) {
    if (![...wanted].some((id) => id.startsWith(`${file.name}#`))) continue
    const units = await loadUnits(file)
    try {
      for (const unit of units) {
        const id = `${file.name}#${unit.index}`
        if (!wanted.has(id)) continue
        const bitmap = await renderUnit(unit, 1)
        if ('skipped' in bitmap) throw new Error(`${id} skipped: ${bitmap.skipped}`)
        pages.push({ id, kind: file.kind, bitmap })
      }
    } finally { units.dispose() }
  }
  const missing = [...wanted].filter((id) => !pages.some((p) => p.id === id))
  if (missing.length) throw new Error(`units not found: ${missing.join(', ')}`)
  const result = await runBattery(pages, { params: [t7.chosen], nullCount: 0, tau: t9.tau })
  const classOf = (id: string) => (DOCX_TOPUP.includes(id) ? 'docx-topup' : 'image-heavy')
  const captures = result.t9.captures as Array<{ page: string; attack: string; cause: string }>
  const survivalByClass = ['image-heavy', 'docx-topup'].map((cls) => ({
    class: cls,
    rows: [...new Set(captures.map((c) => c.attack))].map((attack) => {
      const rows = captures.filter((c) => c.attack === attack && classOf(c.page) === cls)
      const correct = rows.filter((c) => c.cause === 'correct').length
      return { attack, correct, total: rows.length, rate: rows.length ? correct / rows.length : 0 }
    }),
  }))
  const survival = result.t9.survival as Array<{ attack: string; target: number | null; rate: number; status: string }>
  const red = survival.filter((r) => r.status === 'RED')
  writeResult('t9-image-heavy', {
    status: red.length === 0 ? 'GREEN' : 'RED',
    tauFrom: 't9.json (full calibration; not recalibrated)',
    tau: t9.tau,
    params: t7.chosen,
    units: { imageHeavy, docxTopup: DOCX_TOPUP },
    survival,
    survivalByClass,
    collusion: result.t9.collusion,
    captures,
    pageDeltaE: (result.t7.sweep as Array<{ pageMax: number; measurements: unknown[] }>).map((s) => ({ pageMax: s.pageMax, measurements: s.measurements })),
  })
  console.log(`[t9-ih] ${red.length === 0 ? 'GREEN' : 'RED'} ${pages.length} units · ${survival.map((r) => `${r.attack} ${(r.rate * 100).toFixed(1)}%`).join(' · ')}`)
  return red.length === 0 ? 0 : 3
}

if (import.meta.main) process.exit(await main())
