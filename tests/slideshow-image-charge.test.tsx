/**
 * Slideshow plan S1-20 review I1 (fix round 2): the decoded-image charge
 * covers the representation that is actually decoded. codex's case: a
 * picture with a self-contained 1024 x 768 SVG and a 1 x 1 PNG fallback;
 * `decodeImageAsset` prefers the SVG, so the charge must be 1024 x 768 x 4,
 * not the fallback's 4 bytes.
 *
 * jsdom cannot rasterise SVG, so the decoder is replaced (its own file:
 * `vi.mock` is per module) by one that follows the same preference: the SVG
 * candidate when present and accepted, else the primary bytes. Every clock
 * is driven. Decks are "Office-shaped synthetic".
 */
import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import { createCanvas } from 'canvas'
import { runShow } from '../src/slideshow'
import { sessionDeps } from '../src/slideshow/player/session'
import { ManualClock } from '../src/slideshow/clock'
import { createSlideShowWorkspace } from '../src/slideshow/workspace'
import type { ShowStart } from '../src/components/present'
import { buildPptxDeck } from '../src/testdata/ooxml-builders'
import { OfficePackage } from '../src/core/zip'
import { parsePptx, pptxPackageHandle } from '../src/pptx/parse'
import type { PptxDocument } from '../src/pptx/types'

let svgSize = { w: 1024, h: 768 }
vi.mock('../src/core/images', async (importOriginal) => {
  const real = await importOriginal<typeof import('../src/core/images')>()
  return {
    ...real,
    decodeImageAsset: async (asset: { svg?: { verdict: { ok: boolean } }; data?: Uint8Array }) =>
      asset.svg?.verdict.ok ? createCanvas(svgSize.w, svgSize.h) : real.decodeImageAsset(asset as never),
  }
})

const clock = new ManualClock()
const leases: Array<{ tag: string; bytes: number; live: boolean }> = []
let limit: number | undefined
beforeEach(() => {
  leases.length = 0
  limit = undefined
  svgSize = { w: 1024, h: 768 }
  vi.spyOn(sessionDeps, 'clock').mockImplementation(() => clock)
  vi.spyOn(sessionDeps, 'frames').mockImplementation(() => () => {})
  vi.spyOn(sessionDeps, 'workspace').mockImplementation((profile) => {
    const ws = createSlideShowWorkspace(profile, limit !== undefined ? { limit, scratch: 8 * 2 ** 20 } : {})
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
  })
})
afterEach(() => {
  vi.restoreAllMocks()
  document.body.replaceChildren()
})
const flush = async () => { for (let i = 0; i < 20; i++) await new Promise((r) => setTimeout(r, 0)) }

async function svgDeck(): Promise<PptxDocument> {
  const png = new Uint8Array(createCanvas(1, 1).toBuffer('image/png'))
  const doc = await parsePptx(await OfficePackage.load(await buildPptxDeck({
    slides: [{ shapes: [{ id: 2, off: ['0', '0'], ext: ['4572000', '3429000'], image: { data: png } }] }],
  })))
  // The picture's svgBlip candidate: a self-contained 1024 x 768 SVG with the PNG as its raster fallback.
  const svg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg" width="1024" height="768"><rect width="1024" height="768" fill="#369"/></svg>')
  Object.assign(doc.images![0], { svg: { bytes: svg, verdict: { ok: true } } })
  return doc
}

async function show(doc: PptxDocument) {
  const host = document.createElement('div')
  host.setAttribute('data-officeview-show', '')
  document.body.append(host)
  const ac = new AbortController()
  const log: string[] = []
  const s = await runShow({
    doc, handle: pptxPackageHandle(doc), host,
    token: { release() {}, released: false, owns: () => false },
    platform: { hasActivation: () => true, requestFullscreen: async () => {}, exitFullscreen: async () => {}, fullscreenElement: () => null, open: () => null },
    fullscreen: Promise.resolve('entered'), retryFullscreen: async () => 'entered',
    from: 'beginning', motion: 'auto',
    callbacks: { currentSlide: () => 0, restoreFocus() {}, scrollToSlide() {}, onDiagnostic: (d) => log.push(d.code) },
    signal: ac.signal, generation: 1, isCurrent: () => !ac.signal.aborted, cleanup: () => ac.abort(),
  } as ShowStart)
  await flush()
  return { s, log }
}

it('I1 (round 2): an SVG preferred over its raster fallback is charged at its decoded size', async () => {
  const { s } = await show(await svgDeck())
  const image = leases.filter((l) => l.tag.startsWith('image:'))
  expect(image).toEqual([{ tag: 'image:0:0', bytes: 1024 * 768 * 4, live: true }])
  s.exit()
  expect(image[0].live).toBe(false)
})

it('I1 (round 2): a decoded size the workspace cannot hold closes the bitmap and refuses, with a diagnostic', async () => {
  // Room for the reserve, metadata and a small image, not for a 4000 x 4000 decode.
  limit = 1024 * 768 * 4 + 8 * 2 ** 20 + 4 * 2 ** 20
  svgSize = { w: 4000, h: 4000 }
  const { s, log } = await show(await svgDeck())
  expect(leases.filter((l) => l.tag.startsWith('image:') && l.live)).toEqual([])
  expect(log).toContain('image-workspace')
  s.exit()
})
