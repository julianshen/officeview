/**
 * Image-heavy certification (owner ruling 2026-11-08, spec §2 item 14d).
 *
 *   bun run s0:napi scripts/stream-s0/image-heavy.ts
 *
 * Each corpus unit is painted twice at scale 1: once normally, once with
 * `drawImage` disabled on that context. imageShare = pixels that differ /
 * all pixels. It counts every pixel an image changed (transforms and clips
 * included); image pixels identical to what lies beneath are not counted, so
 * the share is a lower bound. A flat-colour picture would pass that alone, so
 * the unit must also be textured: textureShare = fraction of 8×8 blocks whose
 * luminance standard deviation is ≥ 4 (of 255). Image-heavy ⇔ imageShare ≥
 * 0.30 AND textureShare ≥ 0.30.
 */
import { createCanvas } from 'canvas'
import { listCorpus, loadUnits, sizeCeil, type Unit } from './corpus'
import { resolveBackend, writeResult } from './results'

export const IMAGE_HEAVY_SHARE = 0.3
export const TEXTURE_SHARE = 0.3
const BLOCK = 8
const BLOCK_STD = 4

/** Fraction of 8×8 blocks with luminance std ≥ BLOCK_STD. */
export function textureShare(data: Uint8ClampedArray, width: number, height: number): number {
  let blocks = 0
  let textured = 0
  for (let by = 0; by + BLOCK <= height; by += BLOCK) {
    for (let bx = 0; bx + BLOCK <= width; bx += BLOCK) {
      let sum = 0
      let sq = 0
      for (let y = by; y < by + BLOCK; y++) {
        for (let x = bx; x < bx + BLOCK; x++) {
          const i = (y * width + x) * 4
          const l = 0.299 * data[i]! + 0.587 * data[i + 1]! + 0.114 * data[i + 2]!
          sum += l
          sq += l * l
        }
      }
      const n = BLOCK * BLOCK
      const variance = sq / n - (sum / n) ** 2
      blocks++
      if (variance >= BLOCK_STD * BLOCK_STD) textured++
    }
  }
  return blocks > 0 ? textured / blocks : 0
}

function paint(unit: Unit, disableImages: boolean): Uint8ClampedArray {
  const width = sizeCeil(Math.max(1, unit.widthPx))
  const height = sizeCeil(Math.max(1, unit.heightPx))
  const canvas = createCanvas(width, height)
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  if (disableImages) (ctx as { drawImage: unknown }).drawImage = () => {}
  ctx.save()
  ctx.scale(width / unit.widthPx, height / unit.heightPx)
  unit.paint(ctx)
  ctx.restore()
  return ctx.getImageData(0, 0, width, height).data
}

/** Fraction of pixels whose RGBA differs between the two paints. */
export function imageShare(withImages: Uint8ClampedArray, without: Uint8ClampedArray): number {
  let differing = 0
  const pixels = withImages.length / 4
  for (let i = 0; i < withImages.length; i += 4) {
    if (withImages[i] !== without[i] || withImages[i + 1] !== without[i + 1] || withImages[i + 2] !== without[i + 2] || withImages[i + 3] !== without[i + 3]) differing++
  }
  return pixels > 0 ? differing / pixels : 0
}

async function main(): Promise<number> {
  if (resolveBackend() !== '@napi-rs/canvas') {
    console.error('[image-heavy] refused: run with `bun run s0:napi scripts/stream-s0/image-heavy.ts`')
    return 2
  }
  const units: Array<{ id: string; kind: string; width: number; height: number; imageShare: number; textureShare: number; imageHeavy: boolean }> = []
  const failures: Array<{ id: string; reason: string }> = []
  for (const file of listCorpus()) {
    try {
      const loaded = await loadUnits(file)
      try {
        for (const unit of loaded) {
          const id = `${file.name}#${unit.index}`
          if (Math.max(unit.widthPx, unit.heightPx) > 8192) { failures.push({ id, reason: 'oversize' }); continue }
          const full = paint(unit, false)
          const share = imageShare(full, paint(unit, true))
          const texture = textureShare(full, sizeCeil(Math.max(1, unit.widthPx)), sizeCeil(Math.max(1, unit.heightPx)))
          const round = (v: number) => Math.round(v * 10_000) / 10_000
          units.push({ id, kind: file.kind, width: Math.ceil(unit.widthPx), height: Math.ceil(unit.heightPx), imageShare: round(share), textureShare: round(texture), imageHeavy: share >= IMAGE_HEAVY_SHARE && texture >= TEXTURE_SHARE })
        }
      } finally { loaded.dispose() }
    } catch (error) {
      failures.push({ id: file.name, reason: (error as Error).message })
    }
  }
  const certified = units.filter((u) => u.imageHeavy).sort((a, b) => b.imageShare - a.imageShare)
  const counts = Object.fromEntries(['docx', 'xlsx', 'pptx'].map((k) => [k, units.filter((u) => u.kind === k).length]))
  writeResult('image-heavy', {
    status: certified.length >= 5 ? 'GREEN' : 'RED',
    threshold: { imageShare: IMAGE_HEAVY_SHARE, textureShare: TEXTURE_SHARE, blockPx: BLOCK, blockLumaStd: BLOCK_STD },
    minimum: 5,
    method: 'paint twice (drawImage on/off): share of differing pixels (lower bound); plus share of 8x8 blocks with luma std >= 4',
    unitCounts: counts,
    certified: certified.map((u) => ({ id: u.id, imageShare: u.imageShare, textureShare: u.textureShare })),
    units,
    failures,
  })
  console.log(`[image-heavy] ${certified.length} certified ≥ ${IMAGE_HEAVY_SHARE} of ${units.length} units (docx ${counts.docx}, xlsx ${counts.xlsx}, pptx ${counts.pptx}): ${certified.map((u) => `${u.id} ${(u.imageShare * 100).toFixed(1)}%`).join(', ')}`)
  return certified.length >= 5 ? 0 : 3
}

if (import.meta.main) process.exit(await main())
