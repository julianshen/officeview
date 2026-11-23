/**
 * Stage 4a Task 17 — S4 §9 4a "done when", T8 tile pipeline: every corpus XLSX, under each content policy, opens in a
 * tiles-mode host and every visible sheet of it tiles without a static unit. Per sheet: its `sheet-layout` goes through the wire codec
 * (`encodeSheetLayout` → JSON → `decodeSheetLayout`), and tile (0,0) at zoomStep 5, dpr 1 renders 512². Afterwards
 * `paintablesBuilt === 0`: neither `sheetPaintable` nor the too-large refusal paintable ran, so no corpus sheet gets
 * the too-large-sheet notice on the tiles path. A refused layout or tile fails the test. A file the server refuses
 * to open (`unsupported-format`) must be refused identically in units mode, and is named in the log.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { readdirSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { DocumentHost, type OpenResult, type RenderRefusal } from '../src/render'
import type { SheetClock } from '../src/sheet/clock'
import { decodePng } from '../src/mark/png'
import { decodeSheetLayout, encodeSheetLayout, TILE_PX } from '../../src/stream/sheet-wire'

const UTC: SheetClock = { day: '2026-11-23', timeZone: 'UTC' }
const hosts: DocumentHost[] = []
afterAll(() => { for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k) })

describe('sheet corpus (S4 §9 4a, T8)', () => {
  test('every corpus sheet tiles without a static unit', async () => {
    const dir = resolve(import.meta.dir, '../../corpus')
    const names = readdirSync(dir).filter((n) => n.toLowerCase().endsWith('.xlsx')).sort()
    let files = 0, sheets = 0
    const refused: string[] = []
    for (const name of names) {
      const bytes = new Uint8Array(readFileSync(resolve(dir, name)))
      for (const policy of ['visible-only', 'full'] as const) {
        const host = new DocumentHost(); hosts.push(host)
        let o: OpenResult
        try { o = await host.open(name, bytes, policy, { sheetMode: 'tiles' }) } catch (e) {
          // A document the server refuses to open is refused the same way in units mode: the refusal is the
          // projection's or the parser's (it predates 4a), never the tile path's. It is named in the log.
          const units = await new DocumentHost().open(name, bytes, policy).then(() => 'opened', (u: unknown) => String((u as RenderRefusal).code))
          expect({ name, policy, tiles: (e as RenderRefusal).code, units }).toEqual({ name, policy, tiles: 'unsupported-format', units: 'unsupported-format' })
          refused.push(`${name} (${policy})`)
          continue
        }
        expect({ name, policy, sheets: (o.sheets ?? []).length > 0 }).toEqual({ name, policy, sheets: true })
        for (let i = 0; i < o.sheets!.length; i++) {
          const layout = await host.sheetLayout(o.docKey, i, UTC)
          const decoded = decodeSheetLayout(encodeSheetLayout(layout).map((p) => JSON.parse(JSON.stringify({ reqId: 7, ...p }))))
          expect({ name, policy, i, sheetId: decoded.sheetId }).toEqual({ name, policy, i, sheetId: String(i) })
          const tile = await host.renderTile(o.docKey, { sheetIndex: i, viewGen: 0, zoomStep: 5, dpr: 1, palette: 'light', col: 0, row: 0, mark: { kind: 'off' }, clock: UTC })
          const d = decodePng(tile.png)
          expect({ name, policy, i, w: tile.w, h: tile.h, png: [d.w, d.h] }).toEqual({ name, policy, i, w: TILE_PX, h: TILE_PX, png: [TILE_PX, TILE_PX] })
          sheets++
        }
        expect({ name, policy, paintablesBuilt: host.paintablesBuilt }).toEqual({ name, policy, paintablesBuilt: 0 })
        files++
        for (const k of [...(host as unknown as { docs: Map<string, unknown> }).docs.keys()]) host.drop(k)
      }
    }
    expect(files).toBeGreaterThanOrEqual(1)
    // Most of the corpus opens; a refusal is the exception (three files at 13348c79), never the rule.
    expect(refused.length).toBeLessThan(files)
    console.log(`sheet corpus: ${names.length} xlsx; ${files} (file, policy) opens, ${sheets} sheet(s) tiled, paintablesBuilt 0 on every host; refused to open in both modes: ${refused.join(', ') || 'none'}`)
  }, 600_000)
})
