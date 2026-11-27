/**
 * Stage 4a (S4 §6.5, §6.8; R4, R12): the `tiles` request's arguments and the per-sid batch driver. ws.ts owns the
 * socket, the token, the rate, the cache, Layer-1 and the audit; this module validates a request's shape, bounds its
 * coordinates (F8: `tileLimit`) and runs a batch's admitted tiles with at most `perSessionOutstanding` tile jobs in
 * flight per sid, across all of the sid's batches.
 */
import { TILE_BATCH_MAX, TILE_DPRS, TILE_PALETTES, TILE_ZOOM_STEPS, type TilePalette } from '../../../src/stream/sheet-wire'
import { tileLimit } from './tile'

export interface TileCoord { col: number; row: number }
/** A validated `tiles {reqId, token, sheetId, viewGen, zoomStep, dpr, palette, tiles}` (without the token). */
export interface TilesArgs { sheetId: string; sheetIndex: number; viewGen: number; zoomStep: number; dpr: 1 | 2; palette: TilePalette; tiles: TileCoord[] }
/** Per-tile refusal codes in `tiles-done.refused`. `view-refused` (Task 14a fix round 1): the sheet's sort data could
 * not be prepared; its entry carries `reason: 'too-large'`. */
export type TileRefusalCode = 'rate-limited' | 'policy-refused' | 'view-refused'

const SHEET_ID = /^(0|[1-9][0-9]{0,8})$/
const u32 = (n: unknown): n is number => Number.isInteger(n) && (n as number) >= 0 && (n as number) <= 0xffffffff

/**
 * The request's shape, needing no sheet geometry: `sheetId` the canonical decimal index of one of `sheetCount`
 * visible sheets (R12); `zoomStep` an index into `SHEET_ZOOM_STEPS`; `dpr` ∈ {1, 2}; `palette` ∈ {light, dark};
 * `tiles` 1 to 64 distinct `{col, row}` u32 pairs. Anything else is `'policy-refused'`. A well-formed `viewGen` other
 * than the sheet's current one (`viewGenOf`, Stage 4b: the session's; 0 for an identity view never changed) is
 * `'stale-view'`.
 */
export function tilesArgs(msg: Record<string, unknown>, sheetCount: number, viewGenOf: (sheetIndex: number) => number = () => 0): TilesArgs | 'policy-refused' | 'stale-view' {
  const { sheetId, viewGen, zoomStep, dpr, palette, tiles } = msg
  if (typeof sheetId !== 'string' || !SHEET_ID.test(sheetId) || Number(sheetId) >= sheetCount) return 'policy-refused'
  if (!u32(viewGen)) return 'policy-refused'
  if (!Number.isInteger(zoomStep) || (zoomStep as number) < 0 || (zoomStep as number) >= TILE_ZOOM_STEPS.length) return 'policy-refused'
  if (!(TILE_DPRS as readonly unknown[]).includes(dpr)) return 'policy-refused'
  if (!(TILE_PALETTES as readonly unknown[]).includes(palette)) return 'policy-refused'
  if (!Array.isArray(tiles) || tiles.length === 0 || tiles.length > TILE_BATCH_MAX) return 'policy-refused'
  const seen = new Set<string>(), out: TileCoord[] = []
  for (const t of tiles as unknown[]) {
    if (!t || typeof t !== 'object' || Array.isArray(t)) return 'policy-refused'
    const { col, row } = t as Record<string, unknown>
    if (!u32(col) || !u32(row)) return 'policy-refused'
    const k = `${col},${row}`
    if (seen.has(k)) return 'policy-refused'
    seen.add(k); out.push({ col, row })
  }
  if (viewGen !== viewGenOf(Number(sheetId))) return 'stale-view'
  return { sheetId, sheetIndex: Number(sheetId), viewGen, zoomStep: zoomStep as number, dpr: dpr as 1 | 2, palette: palette as TilePalette, tiles: out }
}

/** The sheet's whole domain in logical CSS px (its geometry's `totalCss()` per axis). */
export interface SheetExtent { colsCss: number; rowsCss: number }
/** F8: every coordinate lies on the sheet's `tileRangeFor` grid at `(zoomStep, dpr)`. */
export function tilesInGrid(a: Pick<TilesArgs, 'zoomStep' | 'dpr' | 'tiles'>, e: SheetExtent): boolean {
  const zoom = TILE_ZOOM_STEPS[a.zoomStep]!
  const cols = tileLimit(e.colsCss, zoom, a.dpr), rows = tileLimit(e.rowsCss, zoom, a.dpr)
  return a.tiles.every((t) => t.col < cols && t.row < rows)
}

/** One `tiles` request in progress. `stopped`: `sheet-cancel`, a resume, the session's teardown or (Stage 4b, P14) a
 * view change of its sheet; no tile starts after. `sheetIndex`: set once its arguments are valid. */
export interface TileBatch { reqId: number; stopped: boolean; sheetIndex?: number }

/** R4: a sid's tile jobs in flight (render plus Layer-1), at most `limit`, shared by all of its batches. FIFO. */
export class TileSlots {
  private busy = 0
  private waiters: Array<() => void> = []
  constructor(readonly limit: number) {}
  acquire(): Promise<() => void> {
    return new Promise((resolve) => {
      const grant = () => {
        this.busy++
        let done = false
        resolve(() => { if (done) return; done = true; this.busy--; this.waiters.shift()?.() })
      }
      if (this.busy < this.limit) grant(); else this.waiters.push(grant)
    })
  }
  get inFlight() { return this.busy }
}

/**
 * Runs `tiles` in order: each starts once a slot is free, unless the batch was stopped or `alive` fails by then.
 * `deliver` renders, marks, records and frames one tile and returns its frames or a refusal code; `undefined` means
 * the caller died meanwhile. A tile's frames are sent only if `alive()` holds right then (`send` is synchronous after
 * that check). A tile never throws out of here: any failure is a `policy-refused` entry. Resolves with the refused
 * tiles once every started tile has finished.
 */
export async function runTiles(o: {
  batch: TileBatch; tiles: readonly TileCoord[]; slots: TileSlots; alive: () => boolean
  deliver: (t: TileCoord) => Promise<Uint8Array[][] | TileRefusalCode | undefined>
  send: (frames: Uint8Array[][]) => void
}): Promise<Array<TileCoord & { code: TileRefusalCode }>> {
  const refused: Array<TileCoord & { code: TileRefusalCode }> = []
  const running: Promise<void>[] = []
  for (const t of o.tiles) {
    if (o.batch.stopped || !o.alive()) break
    const release = await o.slots.acquire()
    if (o.batch.stopped || !o.alive()) { release(); break }
    running.push((async () => {
      try {
        const r = await o.deliver(t)
        if (r === undefined) return
        if (typeof r === 'string') { refused.push({ col: t.col, row: t.row, code: r }); return }
        if (o.alive()) o.send(r)
      } catch { refused.push({ col: t.col, row: t.row, code: 'policy-refused' }) } finally { release() }
    })())
  }
  await Promise.all(running)
  return refused
}
