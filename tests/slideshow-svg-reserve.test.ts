/**
 * Slideshow PR #18 review 2, P1: an SVG's bitmap memory is reserved before it is decoded.
 *
 * A picture's accepted svgBlip declares 20000 × 20000 px (1.6 GB as a bitmap). The show never lets a decoder
 * allocate that: it leases the size the picture is painted at (the full rung's device px) and draws the SVG
 * into exactly that. When the SVG fails, the raster fallback is sized from its header and leased before it
 * is decoded too.
 *
 * Both decoders are wrapped (their own file: `vi.mock` is per module) to record the workspace's live image
 * leases at the moment each is called; nothing here rasterises a 20000 px SVG. Every clock is driven.
 * Decks are "Office-shaped synthetic".
 */
import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import { createCanvas } from 'canvas'
import { sessionDeps } from '../src/slideshow/player/session'
import { createSlideShowWorkspace } from '../src/slideshow/workspace'
import { buildPptxDeck } from '../src/testdata/ooxml-builders'
import { installShowSeams, loadDeck, startShow, type ShowSeams } from './helpers/slideshow-show-rig'

const leases: Array<{ tag: string; bytes: number; live: boolean }> = []
const liveImage = () => leases.filter(l => l.live && l.tag.startsWith('image:')).map(l => ({ tag: l.tag, bytes: l.bytes }))
const calls: Array<{ decoder: string; w?: number; h?: number; reserved: Array<{ tag: string; bytes: number }> }> = []
const svgRaster = { fails: false }

vi.mock('../src/core/images', async (importOriginal) => {
  const real = await importOriginal<typeof import('../src/core/images')>()
  return {
    ...real,
    // A decode without the show's reserving decoder would reach the library's own (unreserved) decoders.
    decodeImageAsset: (asset: never, decode?: Parameters<typeof real.decodeImageAsset>[1]) =>
      real.decodeImageAsset(asset, decode ?? (async (_b, mime) => { calls.push({ decoder: `unreserved:${mime}`, reserved: liveImage() }); return undefined })),
    decodeImage: async (bytes: Uint8Array, mime?: string) => {
      calls.push({ decoder: 'raster', reserved: liveImage() })
      return real.decodeImage(bytes, mime)
    },
  }
})
vi.mock('../src/slideshow/render/svg-raster', async (importOriginal) => {
  const real = await importOriginal<typeof import('../src/slideshow/render/svg-raster')>()
  return {
    ...real,
    rasteriseSvg: async (_bytes: Uint8Array, w: number, h: number) => {
      calls.push({ decoder: 'svg', w, h, reserved: liveImage() })
      return svgRaster.fails ? undefined : createCanvas(w, h)
    },
  }
})

let seams: ShowSeams | undefined
const realWorkspace = sessionDeps.workspace
beforeEach(() => {
  leases.length = 0
  calls.length = 0
  svgRaster.fails = false
  sessionDeps.workspace = (profile) => {
    const ws = createSlideShowWorkspace(profile)
    const lease = ws.lease.bind(ws)
    ws.lease = (bytes, tag) => {
      const l = lease(bytes, tag)
      if (!l) return l
      const row = { tag, bytes, live: true }
      leases.push(row)
      const grow = l.grow.bind(l), shrink = l.shrink.bind(l), release = l.release.bind(l)
      l.grow = (b) => { const ok = grow(b); if (ok) row.bytes += b; return ok }
      l.shrink = (b) => { shrink(b); row.bytes -= b }
      l.release = () => { row.live = false; release() }
      return l
    }
    return ws
  }
})
afterEach(() => {
  sessionDeps.workspace = realWorkspace
  seams?.restore()
  seams = undefined
  document.body.replaceChildren()
})

/** A 960 × 720 px slide with one 480 × 360 px picture: a 2 × 1 PNG fallback and a 20000 × 20000 svgBlip. */
async function hugeSvgDeck() {
  const png = new Uint8Array(createCanvas(2, 1).toBuffer('image/png'))
  const doc = await loadDeck(await buildPptxDeck({
    slideSize: [9144000, 6858000],
    slides: [{ shapes: [{ id: 2, off: ['0', '0'], ext: ['4572000', '3429000'], image: { data: png } }] }],
  }))
  const svg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg" width="20000" height="20000"><rect width="20000" height="20000" fill="#369"/></svg>')
  Object.assign(doc.images![0], { svg: { bytes: svg, verdict: { ok: true } } })
  return doc
}

it('an SVG declaring 20000 × 20000 is drawn at its painted size, reserved before it is decoded', async () => {
  seams = installShowSeams({ width: 1280, height: 720 })
  const r = await startShow(await hugeSvgDeck(), { seams, size: { width: 1280, height: 720 } })
  try {
    // 1280 × 720 at DPR 1 shows the 960 × 720 slide at scale 1: the picture paints 480 × 360 device px.
    expect(calls).toEqual([{ decoder: 'svg', w: 480, h: 360, reserved: [{ tag: 'image:0:0', bytes: 480 * 360 * 4 }] }])
    expect(liveImage()).toEqual([{ tag: 'image:0:0', bytes: 480 * 360 * 4 }])
    expect(r.diagnostics.filter(d => d.code === 'image-workspace')).toEqual([])
  } finally { r.exit() }
  expect(liveImage()).toEqual([])
})

it('when the SVG fails, its raster fallback is sized from its header and reserved before it is decoded', async () => {
  svgRaster.fails = true
  seams = installShowSeams({ width: 1280, height: 720 })
  const r = await startShow(await hugeSvgDeck(), { seams, size: { width: 1280, height: 720 } })
  try {
    expect(calls).toEqual([
      { decoder: 'svg', w: 480, h: 360, reserved: [{ tag: 'image:0:0', bytes: 480 * 360 * 4 }] },
      // The SVG's reservation is returned first: only the fallback's 2 × 1 header size is held.
      { decoder: 'raster', reserved: [{ tag: 'image:0:0', bytes: 2 * 1 * 4 }] },
    ])
    expect(liveImage()).toEqual([{ tag: 'image:0:0', bytes: 2 * 1 * 4 }])
  } finally { r.exit() }
})
