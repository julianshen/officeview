/**
 * S2-8 fix round 1: text builds by word wired into the live stage (spec 6.5,
 * 7). jsdom with node-canvas: the real stage runs on the Canvas 2D
 * compositor. The show clock and frames are injected and driven; nothing
 * waits on wall time. Decks are "Office-shaped synthetic".
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { runShow } from '../src/slideshow'
import { sessionDeps, type SessionController } from '../src/slideshow/player/session'
import { ManualClock } from '../src/slideshow/clock'
import { createSlideShowWorkspace } from '../src/slideshow/workspace'
import type { ShowStart } from '../src/components/present'
import { buildPptxDeck, mainSeqXml, timingXmlOf, type PptxDeckSpec, type PptxShapeSpec } from '../src/testdata/ooxml-builders'
import { OfficePackage } from '../src/core/zip'
import { parsePptx, pptxPackageHandle } from '../src/pptx/parse'
import type { PptxDocument } from '../src/pptx/types'
import type { RenderState } from '../src/slideshow/types'

/** A click step whose entrance appears `spid` by word, `ms` apart. */
const byWord = (id: number, spid: number, ms: number) =>
  `<p:par><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="${id + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="${id + 2}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst>` +
  `<p:iterate type="wd"><p:tmAbs val="${ms}"/></p:iterate><p:childTnLst>` +
  `<p:set><p:cBhvr><p:cTn id="${id + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>` +
  `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`

const at = (id: number, x: number, y: number, w: number, h: number, text: string): PptxShapeSpec => ({
  id, name: `Shape ${id}`, off: [String(x * 9525), String(y * 9525)], ext: [String(w * 9525), String(h * 9525)], fill: '336699',
  paragraphs: [{ runs: [{ text }] }],
})

async function deck(spec: PptxDeckSpec): Promise<PptxDocument> {
  return parsePptx(await OfficePackage.load(await buildPptxDeck(spec)))
}

let clock: ManualClock
beforeEach(() => {
  clock = new ManualClock()
  vi.spyOn(sessionDeps, 'clock').mockImplementation(() => clock)
  vi.spyOn(sessionDeps, 'frames').mockImplementation(() => () => {})
  vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (this: HTMLElement) {
    return { left: 0, top: 0, right: 1024, bottom: 768, width: 1024, height: 768, x: 0, y: 0, toJSON() {} } as DOMRect
  })
})
afterEach(() => {
  vi.restoreAllMocks()
  document.body.replaceChildren()
})

const flush = async () => { for (let i = 0; i < 20; i++) await new Promise((r) => setTimeout(r, 0)) }

async function show(doc: PptxDocument): Promise<{ s: SessionController; log: string[]; key(k: string): Promise<void>; tick(ms: number): Promise<void> }> {
  const host = document.createElement('div')
  host.tabIndex = -1
  host.setAttribute('data-officeview-show', '')
  host.style.visibility = 'visible'
  document.body.append(host)
  const ac = new AbortController()
  const log: string[] = []
  const s = await runShow({
    doc, handle: pptxPackageHandle(doc), host,
    token: { release() {}, released: false, owns: () => false },
    platform: { hasActivation: () => true, requestFullscreen: async () => {}, exitFullscreen: async () => {}, fullscreenElement: () => null, open: () => null },
    fullscreen: Promise.resolve('entered'), retryFullscreen: async () => 'entered',
    from: 'beginning', motion: 'auto',
    callbacks: {
      onStart: () => {}, onEnd: () => {}, onSlideChange: () => {},
      onDiagnostic: (d: { code: string }) => log.push(`diag:${d.code}`),
      currentSlide: () => 0, restoreFocus: () => {}, scrollToSlide: () => {},
    },
    signal: ac.signal, generation: 1, isCurrent: () => !ac.signal.aborted,
    cleanup: () => { ac.abort(); host.style.visibility = 'hidden' },
  } as ShowStart)
  await flush()
  s.tick()
  return {
    s, log,
    async key(k) { host.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true })); await flush(); s.tick() },
    async tick(ms) { clock.advance(ms); s.tick(); await flush(); s.tick() },
  }
}

const shown = (state: RenderState, key: string) => state.get(key)?.visible ?? true
const WORDS = ['sp:2/c:0-4', 'sp:2/c:6-9', 'sp:2/c:11-15']

describe('S2-8 live stage', () => {
  it('a by-word entrance reveals the words one at a time', async () => {
    const doc = await deck({ slides: [{ shapes: [at(2, 40, 40, 480, 120, 'Alpha beta gamma')], timingXml: timingXmlOf(mainSeqXml([byWord(10, 2, 500)])) }] })
    const r = await show(doc)
    const now = () => r.s.stage!.stateAt(r.s.cursor, clock.now())
    // Before the click: the body and every word are hidden.
    expect([shown(now(), 'sp:2'), ...WORDS.map(k => shown(now(), k))]).toEqual([false, false, false, false])
    // Every word is its own layer.
    const layers = r.s.stage!.plan!.layers.flatMap(l => l.kind === 'target' ? [`${l.key}:${l.part ?? ''}`] : [])
    expect(layers).toEqual(['sp:2:body', ...WORDS.map(k => `${k}:range`)])
    await r.key('ArrowRight')
    await r.tick(10)
    expect(WORDS.map(k => shown(now(), k))).toEqual([true, false, false])
    expect(shown(now(), 'sp:2')).toBe(true)
    await r.tick(500)
    expect(WORDS.map(k => shown(now(), k))).toEqual([true, true, false])
    await r.tick(500)
    expect(WORDS.map(k => shown(now(), k))).toEqual([true, true, true])
    expect(r.s.stage!.inspect().lastDraws).toEqual(expect.arrayContaining(WORDS.map(k => `target:${k}:range`)))
    expect(r.log).toContain('diag:text-build-approximate')
    r.s.exit()
  })

  it('at the terminal rung a split shape paints whole, diagnosed', async () => {
    vi.spyOn(sessionDeps, 'workspace').mockImplementation(() => createSlideShowWorkspace('desktop', { limit: 1024 * 768 * 4 + 8 * 2 ** 20 + 2 ** 20, scratch: 8 * 2 ** 20 }))
    const doc = await deck({ slides: [{ shapes: [at(2, 0, 0, 480, 360, 'Alpha beta gamma')], timingXml: timingXmlOf(mainSeqXml([byWord(10, 2, 500)])) }] })
    const r = await show(doc)
    expect(r.s.stage!.level).toBe('terminal')
    expect(r.log).toContain('diag:text-build-whole')
    r.s.exit()
  })
})

/** A click step with one whole-shape behaviour on `spid` (an exit or an emphasis). */
const wholeShape = (id: number, cls: 'exit' | 'emph', behaviour: string) =>
  `<p:par><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="${id + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="${id + 2}" presetID="1" presetClass="${cls}" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  behaviour + `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`
const hide = (id: number, spid: number) =>
  `<p:set><p:cBhvr><p:cTn id="${id}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="hidden"/></p:to></p:set>`
const moveRight = (id: number, spid: number) =>
  `<p:animMotion origin="layout" path="M 0 0 L 0.25 0 E" pathEditMode="relative"><p:cBhvr><p:cTn id="${id}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:animMotion>`

describe('PR #19 review 6', () => {
  it('a whole-shape exit after a by-word entrance hides the words too', async () => {
    const doc = await deck({ slides: [{ shapes: [at(2, 40, 40, 480, 120, 'Alpha beta gamma')], timingXml: timingXmlOf(mainSeqXml([byWord(10, 2, 100), wholeShape(20, 'exit', hide(23, 2))])) }] })
    const r = await show(doc)
    const now = () => r.s.stage!.stateAt(r.s.cursor, clock.now())
    await r.key('ArrowRight')
    await r.tick(500)
    expect([shown(now(), 'sp:2'), ...WORDS.map(k => shown(now(), k))]).toEqual([true, true, true, true])
    // The exit targets the whole shape: its words are part of it.
    await r.key('ArrowRight')
    await r.tick(50)
    expect([shown(now(), 'sp:2'), ...WORDS.map(k => shown(now(), k))]).toEqual([false, false, false, false])
    r.s.exit()
  })

  it('a whole-shape emphasis after a by-word entrance moves the words with the shape', async () => {
    const doc = await deck({ slides: [{ shapes: [at(2, 40, 40, 480, 120, 'Alpha beta gamma')], timingXml: timingXmlOf(mainSeqXml([byWord(10, 2, 100), wholeShape(20, 'emph', moveRight(23, 2))])) }] })
    const r = await show(doc)
    const now = () => r.s.stage!.stateAt(r.s.cursor, clock.now())
    await r.key('ArrowRight')
    await r.tick(500)
    await r.key('ArrowRight')
    await r.tick(50)
    const offset = (key: string) => { const s = now().get(key)!; return s.transform[4] + (s.motion?.dx ?? 0) }
    const body = offset('sp:2')
    expect(body).toBeGreaterThan(100)
    for (const k of WORDS) expect(offset(k)).toBeCloseTo(body, 6)
    r.s.exit()
  })
})
