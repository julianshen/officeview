/**
 * Stage-0 Task 2 — node-canvas child runner.
 *
 * The bun preload rewrites `canvas` → the napi shim in the parent process, so
 * node-canvas renders must happen in a **bun child WITHOUT the preload**
 * (bun's module resolution then picks the real `canvas` package).
 *
 * Usage: bun scripts/stream-s0/t1-node-child.ts <file> <unitIndex> <zoom> <dpr> <outPng>
 * Writes the rendered PNG to `<outPng>` and prints a JSON status line.
 */
import { readFileSync } from 'node:fs'
import path from 'node:path'
import { loadOfficeFile } from '../../src/components/OfficeFile'
import { getPaintables } from '../../src/render/paint'
import { createCanvas } from 'canvas'

const [file, unitIndexRaw, zoomRaw, dprRaw, outPng] = process.argv.slice(2)
const unitIndex = Number(unitIndexRaw)
const zoom = Number(zoomRaw)
const dpr = Number(dprRaw)

if (!file || !Number.isInteger(unitIndex) || !Number.isFinite(zoom) || !Number.isFinite(dpr) || !outPng) {
  console.error(JSON.stringify({ ok: false, error: 'usage: t1-node-child <file> <unitIndex> <zoom> <dpr> <outPng>' }))
  process.exit(2)
}

const bytes = readFileSync(file)
const doc = await loadOfficeFile(bytes)
const paintables = await getPaintables(doc)
const paintable = paintables[unitIndex]
if (!paintable) {
  console.error(JSON.stringify({ ok: false, error: `unit index ${unitIndex} out of range (${paintables.length})` }))
  paintables.dispose()
  process.exit(3)
}

const width = Math.max(1, Math.ceil(paintable.spec.widthPx * zoom * dpr))
const height = Math.max(1, Math.ceil(paintable.spec.heightPx * zoom * dpr))
if (Math.max(width, height) > 8192) {
  console.error(JSON.stringify({ ok: false, error: 'oversize' }))
  paintables.dispose()
  process.exit(4)
}

const canvas = createCanvas(width, height)
const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
ctx.save()
ctx.scale(canvas.width / paintable.spec.widthPx, canvas.height / paintable.spec.heightPx)
try {
  paintable.paint(ctx as never)
} finally {
  paintables.dispose()
}
ctx.restore()
const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data

// PNG encode via the node-canvas backend itself.
const png = canvas.toBuffer('image/png')
const { mkdirSync, writeFileSync } = await import('node:fs')
mkdirSync(path.dirname(outPng), { recursive: true })
writeFileSync(outPng, png)
void data
console.log(JSON.stringify({ ok: true, width: canvas.width, height: canvas.height, out: outPng }))
