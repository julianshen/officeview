/**
 * Throwaway Stage-2c spike: bytes on the wire for final-raster vs an estimated
 * culled-commands payload vs relaxed commands. Not product code.
 *
 *   bun --preload ./scripts/stream-s0/napi-preload.ts scripts/stream-s2c-measure/measure.ts [out.json] [fileFilter]
 */
import { readFileSync, readdirSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { gzipSync } from 'node:zlib'
import { DocumentHost, criticalChunksOnly, type PaintObserver, type RenderedUnit } from '../../server/src/render'
import { markRaster } from '../../server/src/mark/mark-raster'
import { deriveBasisKey } from '../../server/src/mark/basis'
import { decodePng } from '../../server/src/mark/png'
import { encodeMessage } from '../../src/stream/frame'
import { encodeUnit } from '../../src/stream/unit'
import { encodeRasterBody } from '../../src/stream/raster'
import { AssetTooLarge, encodeAssets } from '../../src/stream/asset-wire'
import { ASSET_ALLOWANCE } from '../../src/stream/remote/reservation'
import { paintWatermark, setWatermarkMeasurer } from '../../src/core/watermark'
import { markOptions } from '../../server/src/render'

const root = resolve(import.meta.dir, '../..')
const OUT = process.argv[2] ?? resolve(root, 'scripts/stream-s2c-measure/out.json')
const FILTER = process.argv[3]
const KIND = { ASSET_DEFINE: 2, UNIT_RASTER: 3, UNIT_COMMANDS: 4 } as const
const CAP: Record<number, number> = { 2: 1 << 20, 3: 8 << 20, 4: 8 << 20 }
/** Same framing as ws.ts `frame`: gzip level 6, 1 MiB fragments; wire bytes = sum of fragments. */
function frameBytes(kind: number, payload: Uint8Array): number | undefined {
  const z = new Uint8Array(gzipSync(payload, { level: 6 }))
  if (z.length + 4 > CAP[kind]!) return undefined
  return encodeMessage({ ver: 1, kind: kind as never, reqId: 1, payload, compress: () => z, fragmentBytes: 1 << 20 }).reduce((n, f) => n + f.length, 0)
}

const MASTER = new Uint8Array(32).fill(7)
const CODE = (() => { const c = new Int8Array(512); let s = 12345; for (let i = 0; i < 512; i++) { s = (s * 1103515245 + 12345) >>> 0; c[i] = s & 0x10000 ? 1 : -1 } return c })()

/** Forensic-on wire bytes of one raster (marked if markable, else canonical). */
function markedRasterBytes(docRef: string, unitTag: string, png: Uint8Array, w: number, h: number) {
  const r = markRaster({ png, w, h, kBasis: deriveBasisKey(MASTER, docRef, unitTag), code: CODE })
  const out = r.png
  return { png: out, unmarked: 'unmarked' in r ? r.unmarked : undefined }
}
function unitRasterFrame(unitId: number, w: number, h: number, png: Uint8Array) {
  return frameBytes(KIND.UNIT_RASTER, encodeUnit({ unitId, discriminator: 2, body: encodeRasterBody(w, h, png) }))
}
function commandsFrames(unitId: number, unit: RenderedUnit & { kind: 'commands' }): { bytes?: number; why?: string } {
  let payloads: Uint8Array[]
  try { payloads = encodeAssets(unit.assets) } catch (e) { if (e instanceof AssetTooLarge) return { why: 'asset-too-large' }; throw e }
  if (payloads.reduce((n, p) => n + p.length, 0) > ASSET_ALLOWANCE) return { why: 'asset-allowance' }
  let n = 0
  for (const p of payloads) { const f = frameBytes(KIND.ASSET_DEFINE, p); if (f === undefined) return { why: 'asset-frame-cap' }; n += f }
  const f = frameBytes(KIND.UNIT_COMMANDS, encodeUnit({ unitId, discriminator: 1, body: unit.ops }))
  if (f === undefined) return { why: 'commands-frame-cap' }
  return { bytes: n + f }
}

type Variant = 'A' | 'B'
/** Drops content-bearing draw ops; records unmodelled state among KEPT ops (fail-closed proof inputs).
 * A: content = text runs + images. B: A + any fill/stroke whose path has a curve, a Path2D, or a filled non-rect polygon. */
function filterObserver(variant: Variant) {
  const unmodelled: Record<string, number> = {}
  let dropped = 0, kept = 0
  let pathCurve = false, pathPoly = false
  const bump = (k: string) => { unmodelled[k] = (unmodelled[k] ?? 0) + 1 }
  const PAINT = new Set(['fill', 'stroke', 'fillRect', 'strokeRect', 'clearRect', 'putImageData'])
  const obs: PaintObserver & { stats: () => { unmodelled: Record<string, number>; dropped: number; kept: number } } = {
    beforeUnit() {}, afterUnit() {},
    wrap(ctx) {
      return new Proxy(ctx, {
        get(target, key) {
          const v = Reflect.get(target, key, target)
          if (typeof v !== 'function') return v
          return (...args: unknown[]) => {
            if (key === 'fillText' || key === 'strokeText' || key === 'drawImage') { dropped++; return undefined }
            if (key === 'beginPath') { pathCurve = false; pathPoly = false }
            if (key === 'bezierCurveTo' || key === 'quadraticCurveTo' || key === 'arc' || key === 'arcTo' || key === 'ellipse' || key === 'roundRect') pathCurve = true
            if (key === 'lineTo') pathPoly = true
            if (variant === 'B' && (key === 'fill' || key === 'stroke')) {
              const p2d = args.length > 0 && typeof args[0] === 'object'
              if (p2d || pathCurve || (key === 'fill' && pathPoly)) { dropped++; return undefined }
            }
            if (typeof key === 'string' && PAINT.has(key)) {
              kept++
              const t = target as CanvasRenderingContext2D
              if (key === 'putImageData') bump('putImageData')
              if (key === 'clearRect') bump('clearRect')
              if (t.globalCompositeOperation !== 'source-over') bump(`composite:${t.globalCompositeOperation}`)
              if (t.filter && t.filter !== 'none') bump('filter')
              if ((key === 'fill' || key === 'stroke') && args.length > 0 && typeof args[0] === 'object') bump('path2d')
            }
            if (key === 'clip' && args.length > 0 && typeof args[0] === 'object') bump('clip-path2d')
            return Reflect.apply(v, target, args)
          }
        },
        set(target, key, value) { return Reflect.set(target, key, value, target) },
      })
    },
    stats: () => ({ unmodelled, dropped, kept }),
  }
  return obs
}

type Rect = { x: number; y: number; w: number; h: number }
const CELL = 16, MIN_SIDE = 92 // 23 × 23 4-px cells = 529 ≥ 512 Layer-1 chips
/** Content blocks = bounding rects of 8-connected CELL×CELL cells where the full and
 * content-stripped renders differ; overlapping rects merged; each grown to Layer-1 capacity. */
function contentBlocks(a: Uint8ClampedArray, b: Uint8ClampedArray, w: number, h: number): { blocks: Rect[]; diffPx: number } {
  const cw = Math.ceil(w / CELL), ch = Math.ceil(h / CELL), cell = new Uint8Array(cw * ch)
  let diffPx = 0
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const i = (y * w + x) * 4
    if (a[i] !== b[i] || a[i + 1] !== b[i + 1] || a[i + 2] !== b[i + 2] || a[i + 3] !== b[i + 3]) { diffPx++; cell[Math.floor(y / CELL) * cw + Math.floor(x / CELL)] = 1 }
  }
  const seen = new Uint8Array(cw * ch); let rects: Rect[] = []
  for (let s = 0; s < cell.length; s++) {
    if (!cell[s] || seen[s]) continue
    let x0 = Infinity, y0 = Infinity, x1 = -1, y1 = -1; const stack = [s]; seen[s] = 1
    while (stack.length) {
      const c = stack.pop()!, cx = c % cw, cy = (c - cx) / cw
      x0 = Math.min(x0, cx); y0 = Math.min(y0, cy); x1 = Math.max(x1, cx); y1 = Math.max(y1, cy)
      for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
        const nx = cx + dx, ny = cy + dy; if (nx < 0 || ny < 0 || nx >= cw || ny >= ch) continue
        const n = ny * cw + nx; if (cell[n] && !seen[n]) { seen[n] = 1; stack.push(n) }
      }
    }
    rects.push({ x: x0 * CELL, y: y0 * CELL, w: Math.min(w, (x1 + 1) * CELL) - x0 * CELL, h: Math.min(h, (y1 + 1) * CELL) - y0 * CELL })
  }
  const grow = (r: Rect): Rect => {
    const nw = Math.min(w, Math.max(r.w, MIN_SIDE)), nh = Math.min(h, Math.max(r.h, MIN_SIDE))
    const nx = Math.max(0, Math.min(w - nw, r.x - Math.floor((nw - r.w) / 2))), ny = Math.max(0, Math.min(h - nh, r.y - Math.floor((nh - r.h) / 2)))
    return { x: nx, y: ny, w: nw, h: nh }
  }
  const ov = (p: Rect, q: Rect) => p.x < q.x + q.w && q.x < p.x + p.w && p.y < q.y + q.h && q.y < p.y + p.h
  const merge = (list: Rect[]) => {
    let changed = true
    while (changed) {
      changed = false
      outer: for (let i = 0; i < list.length; i++) for (let j = i + 1; j < list.length; j++) if (ov(list[i]!, list[j]!)) {
        const p = list[i]!, q = list[j]!, x = Math.min(p.x, q.x), y = Math.min(p.y, q.y)
        list[i] = { x, y, w: Math.max(p.x + p.w, q.x + q.w) - x, h: Math.max(p.y + p.h, q.y + q.h) - y }
        list.splice(j, 1); changed = true; break outer
      }
    }
    return list
  }
  rects = merge(rects)
  for (let k = 0; k < 4; k++) { rects = merge(rects.map(grow)) }
  return { blocks: rects, diffPx }
}

function cropPng(host: DocumentHost, rgba: Uint8ClampedArray, w: number, r: Rect): Uint8Array {
  const c = host.backend.createCanvas(r.w, r.h), ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D
  const img = ctx.createImageData(r.w, r.h)
  for (let y = 0; y < r.h; y++) img.data.set(rgba.subarray(((r.y + y) * w + r.x) * 4, ((r.y + y) * w + r.x + r.w) * 4), y * r.w * 4)
  ctx.putImageData(img, 0, 0)
  return criticalChunksOnly(new Uint8Array((c as unknown as { toBuffer(m: 'image/png'): Buffer }).toBuffer('image/png')))
}

/** Visible-mark overlay alone on a transparent surface: as a raster, and as commands. */
function markOverlayBytes(host: DocumentHost, page: { widthPx: number; heightPx: number }, w: number, h: number, scale: number) {
  const c = host.backend.createCanvas(w, h), ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D
  ctx.scale(scale, scale); setWatermarkMeasurer(null); paintWatermark(ctx, page, markOptions({ kind: 'fixed' })!)
  const png = criticalChunksOnly(new Uint8Array((c as unknown as { toBuffer(m: 'image/png'): Buffer }).toBuffer('image/png')))
  return { pngFrame: frameBytes(KIND.UNIT_RASTER, encodeRasterBody(w, h, png))!, pngBytes: png.length }
}

interface UnitRow {
  file: string; format: string; unit: number; scale: number; w: number; h: number
  finalRaster: number; finalNoVisibleMark?: number; units: number; finalUnmarked?: string
  relaxed: { bytes: number; kind: string; fallback?: string }
  culled: Record<Variant, { qualifies: boolean; reason?: string; commands?: number; rasterBlocks?: number; blocks?: number; blockArea?: number; diffPx?: number; overlay?: number; total?: number; lean?: number; leanHull?: number; hullArea?: number; commandsWithMark?: number; unmarkedBlocks?: number; dropped?: number; kept?: number; unmodelled?: Record<string, number> }>
  ms: number
}

const host = new DocumentHost()
const files = readdirSync(resolve(root, 'corpus')).filter((f) => /\.(pptx|docx|xlsx)$/.test(f) && (!FILTER || f.includes(FILTER))).sort()
const rows: UnitRow[] = []
const docErrors: Array<{ file: string; error: string }> = []
const MAX_UNITS_PER_DOC = Number(process.env.S2C_MAX_UNITS ?? 6)
const t0 = Date.now()
for (const file of files) {
  const format = file.split('.').pop()!
  let opened
  try { opened = await host.open(file, new Uint8Array(readFileSync(resolve(root, 'corpus', file))), 'visible-only') } catch (e) { docErrors.push({ file, error: String((e as Error).message ?? e) }); continue }
  const n = opened.units.length
  // Sampling: first MAX_UNITS_PER_DOC units of each document.
  const units = Array.from({ length: Math.min(n, MAX_UNITS_PER_DOC) }, (_, i) => i)
  const scales = format === 'pptx' ? [1, 2] : [1]
  for (const unit of units) for (const scale of scales) {
    const ts = Date.now()
    try {
      const full = host.render(opened.docKey, { unitIndex: unit, scale, projection: 'final-raster', mark: { kind: 'fixed' } })
      if (full.kind !== 'raster') throw new Error('expected raster')
      const { w, h } = full
      const tag = `${unit}@${scale}`
      const fm = markedRasterBytes(file, tag, full.png, w, h)
      const finalRaster = unitRasterFrame(unit, w, h, fm.png)
      if (finalRaster === undefined) throw new Error('final raster over frame cap')
      // Decomposition only: the same unit with the visible mark off (forensic still on).
      const nv = host.render(opened.docKey, { unitIndex: unit, scale, projection: 'final-raster', mark: { kind: 'off' } })
      const finalNoVisibleMark = nv.kind === 'raster' ? unitRasterFrame(unit, w, h, markedRasterBytes(file, tag, nv.png, w, h).png) : undefined
      // Relaxed commands (real path; falls back to a raster, forensic-marked, when the recorder/faces/asset caps refuse).
      const rel = host.render(opened.docKey, { unitIndex: unit, scale, projection: 'commands', mark: { kind: 'fixed' } })
      let relaxed: UnitRow['relaxed']
      if (rel.kind === 'commands') {
        const c = commandsFrames(unit, rel)
        relaxed = c.bytes !== undefined ? { bytes: c.bytes, kind: 'commands' } : { bytes: finalRaster, kind: 'raster', fallback: c.why }
      } else relaxed = { bytes: finalRaster, kind: 'raster', fallback: rel.fallback }
      const fullRgba = decodePng(full.png).rgba
      const markedRgba = decodePng(fm.png).rgba
      const page = { widthPx: (host as any).docs.get(opened.docKey).raw[unit].widthPx, heightPx: (host as any).docs.get(opened.docKey).raw[unit].heightPx }
      const overlay = markOverlayBytes(host, page, w, h, scale)
      const culled = {} as UnitRow['culled']
      for (const v of ['A', 'B'] as Variant[]) {
        const fo = filterObserver(v)
        const stripped = host.render(opened.docKey, { unitIndex: unit, scale, projection: 'final-raster', mark: { kind: 'fixed' } }, [fo])
        if (stripped.kind !== 'raster') throw new Error('expected raster')
        const st = fo.stats()
        const fc = filterObserver(v)
        const cmd = host.render(opened.docKey, { unitIndex: unit, scale, projection: 'commands', mark: { kind: 'off' } }, [fc])
        const { blocks, diffPx } = contentBlocks(fullRgba, decodePng(stripped.png).rgba, w, h)
        const base = { dropped: st.dropped, kept: st.kept, unmodelled: st.unmodelled, diffPx, blocks: blocks.length }
        const unmodelledKeys = Object.keys(st.unmodelled)
        if (cmd.kind !== 'commands') { culled[v] = { qualifies: false, reason: `recorder:${cmd.fallback}`, ...base }; continue }
        if (unmodelledKeys.length) { culled[v] = { qualifies: false, reason: `unmodelled:${unmodelledKeys.join('+')}`, ...base }; continue }
        const c = commandsFrames(unit, cmd)
        if (c.bytes === undefined) { culled[v] = { qualifies: false, reason: `commands-cap:${c.why}`, ...base }; continue }
        let rasterBlocks = 0, blockArea = 0, unmarkedBlocks = 0
        // Raster blocks are cut from the CANONICAL final composite (visible mark baked in), then Layer-1 marked per block.
        blocks.forEach((r, bi) => {
          const png = cropPng(host, fullRgba, w, r)
          const m = markedRasterBytes(file, `${tag}#b${bi}`, png, r.w, r.h)
          if (m.unmarked) unmarkedBlocks++
          rasterBlocks += frameBytes(KIND.UNIT_RASTER, encodeRasterBody(r.w, r.h, m.png))! + 16
          blockArea += r.w * r.h
        })
        void markedRgba
        // Bracket: one hull block over all content (the max any merge-based segmentation reaches).
        let hullRaster = 0, hullArea = 0
        if (blocks.length) {
          const x0 = Math.min(...blocks.map((b) => b.x)), y0 = Math.min(...blocks.map((b) => b.y))
          const hr = { x: x0, y: y0, w: Math.max(...blocks.map((b) => b.x + b.w)) - x0, h: Math.max(...blocks.map((b) => b.y + b.h)) - y0 }
          const hm = markedRasterBytes(file, `${tag}#hull`, cropPng(host, fullRgba, w, hr), hr.w, hr.h)
          hullRaster = frameBytes(KIND.UNIT_RASTER, encodeRasterBody(hr.w, hr.h, hm.png))! + 16; hullArea = hr.w * hr.h
        }
        const total = c.bytes + rasterBlocks + overlay.pngFrame
        // Lean variant: the visible mark rides the command stream (as relaxed sessions do) instead of a raster overlay.
        const fl = filterObserver(v)
        const cmdM = host.render(opened.docKey, { unitIndex: unit, scale, projection: 'commands', mark: { kind: 'fixed' } }, [fl])
        const cm = cmdM.kind === 'commands' ? commandsFrames(unit, cmdM) : {}
        const lean = cm.bytes !== undefined ? cm.bytes + rasterBlocks : undefined
        const leanHull = cm.bytes !== undefined ? cm.bytes + hullRaster : undefined
        culled[v] = { qualifies: unmarkedBlocks === 0, ...(unmarkedBlocks ? { reason: 'block-unmarkable' } : {}), commands: c.bytes, commandsWithMark: cm.bytes, rasterBlocks, blockArea, overlay: overlay.pngFrame, total, lean, leanHull, hullArea, unmarkedBlocks, ...base }
      }
      rows.push({ file, format, unit, scale, w, h, finalRaster, finalNoVisibleMark, units: n, ...(fm.unmarked ? { finalUnmarked: fm.unmarked } : {}), relaxed, culled, ms: Date.now() - ts })
    } catch (e) {
      docErrors.push({ file: `${file}#${unit}@${scale}`, error: String((e as Error).message ?? e) })
    }
  }
  host.drop(opened.docKey)
  console.error(`${file}: ${n} units (${units.length} sampled) ${((Date.now() - t0) / 1000).toFixed(1)}s`)
  writeFileSync(OUT, JSON.stringify({ rows, docErrors, maxUnitsPerDoc: MAX_UNITS_PER_DOC }, null, 1))
}
writeFileSync(OUT, JSON.stringify({ rows, docErrors, maxUnitsPerDoc: MAX_UNITS_PER_DOC }, null, 1))
console.error(`done ${rows.length} rows, ${docErrors.length} errors, ${((Date.now() - t0) / 1000).toFixed(1)}s`)
