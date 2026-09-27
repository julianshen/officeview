/**
 * Golden-image CLI for any office file.
 *
 *   bun scripts/golden.ts record <file.docx|docx|xlsx|pptx> [--name NAME] [--all]
 *   bun scripts/golden.ts compare <file> [--name NAME] [--threshold N] [--max-ratio 0.01]
 *
 * record: renders every page/sheet/slide to tests/goldens/<name>.unit<i>.png
 * compare: re-renders and diffs against the goldens; writes
 *          /tmp/officeview-diff/<name>.actual.png and .diff.png, exits 1 on mismatch.
 * --all: all units in one tall bitmap per file (single golden image).
 */
import { loadOfficeFile } from '../src/components/OfficeFile'
import { getPaintables } from '../src/render/paint'
import { diffBitmaps, savePng, loadPng, renderPaintables } from '../src/test/pixel-diff'
import { readFileSync, existsSync, mkdirSync } from 'fs'
import { basename, join } from 'path'

const GOLDEN_DIR = join(import.meta.dir, '..', 'tests', 'goldens')

function parseArgs(argv: string[]) {
  const [cmd, file, ...rest] = argv
  const args: Record<string, string | undefined> = {}
  for (let i = 0; i < rest.length; i++) {
    if (rest[i].startsWith('--')) args[rest[i].slice(2)] = rest[i + 1]?.startsWith('--') ? undefined : rest[i + 1]
  }
  return { cmd, file, args }
}

function defaultName(file: string): string {
  return basename(file).replace(/\.(docx|xlsx|pptx)$/i, '')
}

async function main(): Promise<number> {
  const { cmd, file, args } = parseArgs(process.argv.slice(2))
  if (!cmd || !file) {
    console.error('usage: bun scripts/golden.ts record|compare <file> [--name NAME] [--threshold N] [--max-ratio R] [--all]')
    return 2
  }
  const name = args.name ?? defaultName(file)
  const threshold = args.threshold ? parseInt(args.threshold, 10) : 8
  const maxRatio = args['max-ratio'] ? parseFloat(args['max-ratio']) : 0.01
  const doc = await loadOfficeFile(readFileSync(file))
  const paintables = await getPaintables(doc)

  // Render. --all stacks every unit vertically into one bitmap.
  const bitmaps = await renderPaintables(paintables)
  let rendered: Array<{ suffix: string; bm: (typeof bitmaps)[number] }>
  if (args.all) {
    const w = Math.max(...bitmaps.map((b) => b.width))
    const h = bitmaps.reduce((a, b) => a + b.height, 0)
    const { createCanvas } = await import('canvas')
    const canvas = createCanvas(w, h)
    const ctx = canvas.getContext('2d')!
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, w, h)
    let oy = 0
    for (const bm of bitmaps) {
      const tmp = createCanvas(bm.width, bm.height)
      tmp.getContext('2d').putImageData(tmpImageData(bm), 0, 0)
      ctx.drawImage(tmp as never, 0, oy)
      oy += bm.height
    }
    rendered = [{ suffix: 'all', bm: { width: w, height: h, data: ctx.getImageData(0, 0, w, h).data } }]
  } else {
    rendered = bitmaps.map((bm, i) => ({ suffix: String(i), bm }))
  }

  if (cmd === 'record') {
    mkdirSync(GOLDEN_DIR, { recursive: true })
    for (const { suffix, bm } of rendered) {
      const path = join(GOLDEN_DIR, `${name}.${suffix}.png`)
      await savePng(bm, path)
      console.log(`recorded ${path} (${bm.width}x${bm.height})`)
    }
    return 0
  }

  if (cmd === 'compare') {
    let failures = 0
    const diffDir = '/tmp/officeview-diff'
    mkdirSync(diffDir, { recursive: true })
    for (const { suffix, bm } of rendered) {
      const goldenPath = join(GOLDEN_DIR, `${name}.${suffix}.png`)
      if (!existsSync(goldenPath)) {
        console.error(`MISSING golden ${goldenPath} — run record first`)
        failures++
        continue
      }
      const golden = await loadPng(goldenPath)
      const d = diffBitmaps(golden, bm, threshold)
      const pct = (d.ratio * 100).toFixed(3)
      if (d.ratio > maxRatio) {
        failures++
        await savePng(bm, join(diffDir, `${name}.${suffix}.actual.png`))
        await savePng({ width: d.width, height: d.height, data: d.mask } as import('../src/test/pixel-diff').Bitmap, join(diffDir, `${name}.${suffix}.diff.png`))
        console.error(`FAIL ${name}.${suffix}: ${(d.changedPixels).toLocaleString()}/${d.totalPixels.toLocaleString()} px changed (${pct}% > ${(maxRatio * 100).toFixed(2)}%) — actual+diff PNGs in ${diffDir}`)
      } else {
        console.log(`PASS ${name}.${suffix}: ${pct}% changed (maxDelta ${d.maxDelta})`)
      }
    }
    return failures > 0 ? 1 : 0
  }

  console.error(`unknown command: ${cmd}`)
  return 2
}

function tmpImageData(bm: { width: number; height: number; data: Uint8ClampedArray }) {
  const { createImageData } = require('canvas') as { createImageData: (d: Uint8ClampedArray, w: number, h: number) => import('canvas').ImageData }
  return createImageData(bm.data, bm.width, bm.height)
}

process.exit(await main())
