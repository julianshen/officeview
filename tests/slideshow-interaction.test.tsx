/**
 * Slideshow plan S1-20: engine wiring, interactions and accessibility
 * (spec 4.3 click targets, hyperlinks, cursor; 4.4 advClick and kiosk; 4.5
 * reduced motion; 6.3 resize/DPR, watermark and protection; 9).
 *
 * jsdom with node-canvas: the real stage runs on the Canvas 2D compositor
 * (jsdom has no WebGL2, so the backend falls back, diagnosed). The show
 * clock, the frame driver and the wall timers (cursor, restore listener)
 * are all injected and driven; nothing waits on wall time. Decks are
 * "Office-shaped synthetic".
 */
import { readFileSync } from 'node:fs'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { runShow } from '../src/slideshow'
import { RESTORE_LISTENER_MS, sessionDeps, type SessionController } from '../src/slideshow/player/session'
import { hitTest } from '../src/slideshow/player/hit-test'
import { runLink } from '../src/slideshow/player/links'
import { CURSOR_HIDE_MS } from '../src/slideshow/player/cursor'
import { ManualClock } from '../src/slideshow/clock'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { buildRenderPlan } from '../src/slideshow/render/plan'
import { REDUCED_TRANSITION_MS } from '../src/slideshow/render/transitions2d'
import { createSlideShowWorkspace } from '../src/slideshow/workspace'
import { SlideShowError, type ShowStart, type SlideShowCallbacks } from '../src/components/present'
import { buildPptxDeck, mainSeqXml, timingXmlOf, type PptxDeckSpec, type PptxShapeSpec } from '../src/testdata/ooxml-builders'
import { OfficePackage } from '../src/core/zip'
import { parsePptx, pptxPackageHandle } from '../src/pptx/parse'
import type { PptxDocument } from '../src/pptx/types'
import type { LinkSpec, RenderState, TargetState } from '../src/slideshow/types'

// ---------------------------------------------------------------------------
// Timing XML
// ---------------------------------------------------------------------------

/** One effect: a `set` of style.visibility (plus optional extra behaviours) on `spid`, ids `id`, `id + 1`. */
const effect = (id: number, spid: number, cls: 'entr' | 'exit', extra = '') =>
  `<p:par><p:cTn id="${id}" presetID="1" presetClass="${cls}" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:set><p:cBhvr><p:cTn id="${id + 1}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${cls === 'entr' ? 'visible' : 'hidden'}"/></p:to></p:set>` +
  `${extra}</p:childTnLst></p:cTn></p:par>`
/** A 1000 ms fade-in on `spid` with id `id`. */
const fade = (id: number, spid: number) =>
  `<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="${id}" dur="1000"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:animEffect>`
/** A click step holding `effects`, ids `id`, `id + 1`. */
const step = (id: number, effects: string) =>
  `<p:par><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="${id + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${effects}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`
/** An interactive sequence started by a click on `trigger`, ids from `id`. */
const interactive = (id: number, trigger: number, effects: string) =>
  `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${id}" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${trigger}"/></p:tgtEl></p:cond></p:stCondLst><p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync><p:childTnLst>` +
  `<p:par><p:cTn id="${id + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${effects}</p:childTnLst></p:cTn></p:par>` +
  `</p:childTnLst></p:cTn><p:nextCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${trigger}"/></p:tgtEl></p:cond></p:nextCondLst></p:seq>`

/** A box at slide px (x, y), size (w, h), with text. 9525 EMU per px. */
const at = (id: number, x: number, y: number, w: number, h: number, text?: string, extra: Partial<PptxShapeSpec> = {}): PptxShapeSpec => ({
  id, name: `Shape ${id}`, off: [String(x * 9525), String(y * 9525)], ext: [String(w * 9525), String(h * 9525)], fill: '336699',
  ...(text ? { paragraphs: [{ runs: [{ text }] }] } : {}), ...extra,
})

async function deck(spec: PptxDeckSpec): Promise<PptxDocument> {
  return parsePptx(await OfficePackage.load(await buildPptxDeck(spec)))
}

// ---------------------------------------------------------------------------
// Rig: a driven clock, frames and wall timers
// ---------------------------------------------------------------------------

let clock: ManualClock
let timers: Array<{ at: number; cb: () => void; id: number }>
let wall = 0
let nextTimer = 1
const advanceWall = (ms: number) => {
  wall += ms
  for (const t of [...timers].sort((a, b) => a.at - b.at)) if (t.at <= wall && timers.includes(t)) { timers.splice(timers.indexOf(t), 1); t.cb() }
}

beforeEach(() => {
  clock = new ManualClock()
  timers = []
  wall = 0
  vi.spyOn(sessionDeps, 'clock').mockImplementation(() => clock)
  vi.spyOn(sessionDeps, 'frames').mockImplementation(() => () => {})
  vi.spyOn(sessionDeps, 'timer').mockImplementation(() => ({
    set: (cb, ms) => { const id = nextTimer++; timers.push({ at: wall + ms, cb, id }); return id },
    clear: (id) => { const i = timers.findIndex(t => t.id === id); if (i >= 0) timers.splice(i, 1) },
  }))
  vi.spyOn(sessionDeps, 'now').mockImplementation(() => wall)
  // jsdom has no layout: the surface fills a 1024 x 768 viewport (the backing store, DPR 1).
  vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (this: HTMLElement) {
    return { left: 0, top: 0, right: 1024, bottom: 768, width: 1024, height: 768, x: 0, y: 0, toJSON() {} } as DOMRect
  })
})
afterEach(() => {
  vi.restoreAllMocks()
  sessionDeps.backend = undefined
  document.body.replaceChildren()
})

const flush = async () => { for (let i = 0; i < 20; i++) await new Promise((r) => setTimeout(r, 0)) }
/** Slide px → client px: a 960 x 720 slide fills the 1024 x 768 backing store exactly. */
const K = 1024 / 960

interface Rig {
  s: SessionController
  host: HTMLElement
  log: string[]
  opened: Array<[string, string]>
  surface(): HTMLElement
  click(x: number, y: number): Promise<void>
  key(k: string, target?: Element): Promise<void>
  tick(ms?: number): Promise<void>
  text(): string[]
  buttons(): HTMLButtonElement[]
}

function hostIn(parent: Element = document.body): HTMLElement {
  const host = document.createElement('div')
  host.tabIndex = -1
  host.setAttribute('data-officeview-show', '')
  host.style.visibility = 'visible'
  parent.append(host)
  return host
}

async function show(doc: PptxDocument, over: Partial<ShowStart> & { watermark?: { text: string } } = {}, cb: Partial<SlideShowCallbacks> = {}): Promise<Rig> {
  const host = over.host ?? hostIn()
  const ac = new AbortController()
  const log: string[] = []
  const opened: Array<[string, string]> = []
  const start = {
    doc, handle: pptxPackageHandle(doc), host,
    token: { release() {}, released: false, owns: () => false },
    platform: {
      hasActivation: () => true, requestFullscreen: async () => {}, exitFullscreen: async () => {}, fullscreenElement: () => null,
      open: (_w: Window, url: string, features: string) => { opened.push([url, features]); return null },
    },
    fullscreen: Promise.resolve('entered' as const),
    retryFullscreen: async () => 'entered' as const,
    from: 'beginning' as const, motion: 'auto' as const,
    callbacks: {
      onStart: () => log.push('onStart'), onEnd: () => log.push('onEnd'),
      onSlideChange: (i: number, b: number) => log.push(`change:${i}:${b}`),
      onDiagnostic: (d: { code: string }) => log.push(`diag:${d.code}`),
      currentSlide: () => 0, restoreFocus: () => log.push('restoreFocus'), scrollToSlide: (i: number) => log.push(`scroll:${i}`),
      ...cb,
    },
    signal: ac.signal, generation: 1, isCurrent: () => !ac.signal.aborted,
    cleanup: () => { ac.abort(); host.style.visibility = 'hidden' },
    ...over,
  } as ShowStart
  const s = await runShow(start)
  await flush()
  s.tick()
  const surface = () => host.querySelector<HTMLElement>('[data-officeview-show-surface]')!
  return {
    s, host, log, opened, surface,
    async click(x, y) { surface().dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: x * K, clientY: y * K })); await flush(); s.tick() },
    async key(k, target = host) { target.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true })); await flush(); s.tick() },
    async tick(ms = 0) { clock.advance(ms); s.tick(); await flush(); s.tick() },
    text: () => [...host.querySelectorAll('[data-officeview-show-text] li')].map(li => li.textContent || li.getAttribute('aria-label') || ''),
    buttons: () => [...host.querySelectorAll<HTMLButtonElement>('[data-officeview-show-targets] button')],
  }
}

const visible = (state: RenderState, key: string) => state.get(key)?.visible ?? true
const stateNow = (r: Rig) => r.s.stage!.stateAt(r.s.cursor, clock.now())

// ---------------------------------------------------------------------------

describe('click targets', () => {
  // Shape 3 (top-left, 0..96 px) triggers shape 2's entrance; shape 4's entrance is the main sequence's one click step.
  const triggerDeck = () => deck({
    slides: [{
      shapes: [at(2, 400, 300, 200, 100, 'Revealed'), at(3, 0, 0, 96, 96, 'Press'), at(4, 400, 500, 200, 100, 'Main')],
      timingXml: timingXmlOf(mainSeqXml([step(10, effect(12, 4, 'entr'))]) + interactive(20, 3, effect(22, 2, 'entr'))),
    }, { shapes: [] }],
  })

  it('a click on a trigger runs its interactive sequence instead of advancing', async () => {
    const r = await show(await triggerDeck())
    expect(r.text()).toEqual(['Press'])
    expect(visible(stateNow(r), 'sp:2')).toBe(false)
    await r.click(48, 48)
    await r.tick(10)
    expect([r.s.slideIndex, r.s.buildIndex]).toEqual([0, 0])
    expect(r.s.cursor!.events.some(e => e.kind === 'trigger' && e.spid === '3')).toBe(true)
    expect(visible(stateNow(r), 'sp:2')).toBe(true)
    expect(r.text()).toEqual(['Revealed', 'Press'])
    // a click elsewhere advances the main sequence
    await r.click(900, 700)
    expect(r.s.buildIndex).toBe(1)
    r.s.exit()
  })

  it('a click on an exited trigger shape advances instead of firing', async () => {
    // The main step exits the trigger shape itself.
    const doc = await deck({
      slides: [{
        shapes: [at(2, 400, 300, 200, 100, 'Revealed'), at(3, 0, 0, 96, 96, 'Press')],
        timingXml: timingXmlOf(mainSeqXml([step(10, effect(12, 3, 'exit'))]) + interactive(20, 3, effect(22, 2, 'entr'))),
      }, { shapes: [] }],
    })
    const r = await show(doc)
    await r.key('ArrowRight')
    await r.tick(10)
    expect(r.s.buildIndex).toBe(1)
    expect(visible(stateNow(r), 'sp:3')).toBe(false)
    await r.click(48, 48)
    expect(r.s.slideIndex).toBe(1)
    expect(r.log).not.toContain('diag:link-refused')
    r.s.exit()
  })

  it('a shape with both trigger and link runs the trigger', async () => {
    const doc = await deck({
      slides: [{
        shapes: [at(2, 400, 300, 200, 100, 'Revealed'), at(3, 0, 0, 96, 96, 'Press')],
        timingXml: timingXmlOf(interactive(20, 3, effect(22, 2, 'entr'))),
        links: [{ shapeId: 3, target: 'https://example.com/', external: true }],
      }, { shapes: [] }],
    })
    const r = await show(doc)
    await r.click(48, 48)
    await r.tick(10)
    expect(r.opened).toEqual([])
    expect(visible(stateNow(r), 'sp:2')).toBe(true)
    expect(r.s.slideIndex).toBe(0)
    // the accessibility layer lists the shape once, as a trigger
    expect(r.buttons().map(b => b.textContent)).toEqual(['Trigger: Shape 3'])
    r.s.exit()
  })

  it('hit testing walks top-down with transform, mask and visibility', async () => {
    const doc = await deck({ slides: [{ shapes: [at(2, 0, 0, 200, 200), at(3, 100, 100, 200, 200)] }] })
    const slide = doc.slides[0]
    const plan = buildRenderPlan(slide, new Set(['sp:2', 'sp:3']), new Map(), [])
    const I = [1, 0, 0, 1, 0, 0] as const
    const st = (o: Partial<TargetState>): TargetState => ({ visible: true, opacity: 1, transform: I, ...o })
    const both = new Set(['2', '3'])
    expect(hitTest(plan, new Map(), 150, 150, both)).toBe('3')
    expect(hitTest(plan, new Map([['sp:3', st({ visible: false })]]), 150, 150, both)).toBe('2')
    // shape 3 moved 300 px right: its old place now hits shape 2
    expect(hitTest(plan, new Map([['sp:3', st({ motion: { dx: 300, dy: 0 } })]]), 150, 150, both)).toBe('2')
    expect(hitTest(plan, new Map([['sp:3', st({ motion: { dx: 300, dy: 0 } })]]), 450, 150, both)).toBe('3')
    // a wipe from the left at 25 %: the right of shape 3 is clipped away
    const wipe = st({ mask: { kind: 'wipe', progress: 0.25, params: { from: 'l' } } })
    expect(hitTest(plan, new Map([['sp:3', wipe]]), 280, 280, both)).toBeUndefined()
    expect(hitTest(plan, new Map([['sp:3', wipe]]), 120, 280, both)).toBe('3')
    // a non-interactive shape never blocks one below it
    expect(hitTest(plan, new Map(), 150, 150, new Set(['2']))).toBe('2')
  })
})

describe('hyperlinks', () => {
  it('hyperlink slide jump, endshow and url with noopener; refused schemes diagnose', async () => {
    const shapes = [at(2, 0, 0, 96, 96, 'Jump'), at(3, 100, 0, 96, 96, 'End'), at(4, 200, 0, 96, 96, 'Web'), at(5, 300, 0, 96, 96, 'Script')]
    const doc = await deck({
      slides: [{
        shapes,
        links: [
          { shapeId: 2, target: 'slide3.xml', action: 'ppaction://hlinksldjump' },
          { shapeId: 3, action: 'ppaction://hlinkshowjump?jump=endshow' },
          { shapeId: 4, target: 'https://example.com/a', external: true },
          { shapeId: 5, target: 'javascript:alert(1)', external: true },
        ],
      }, { shapes: [] }, { shapes: [] }],
    })
    const r = await show(doc)
    await r.click(48, 48)
    expect(r.s.slideIndex).toBe(2)
    r.s.goTo(0)
    await flush()
    r.s.tick()
    await r.click(248, 48)
    expect(r.opened).toEqual([['https://example.com/a', 'noopener']])
    expect(r.s.slideIndex).toBe(0)
    await r.click(348, 48)
    expect(r.log).toContain('diag:link-refused')
    expect(r.s.slideIndex).toBe(0)
    await r.click(148, 48)
    await r.s.ended
    expect(r.log).toContain('onEnd')

    // runLink on its own: every verb, and a caller-built URL with an unsafe scheme is still refused
    const calls: string[] = []
    const nav = { goTo: (i: number) => calls.push(`goTo:${i}`), next: () => calls.push('next'), previous: () => calls.push('previous'), first: () => calls.push('first'), last: () => calls.push('last'), exit: () => calls.push('exit') }
    const sink = createDiagnosticSink()
    const open = vi.fn(() => null)
    let n = 10
    const link = (action: LinkSpec['action']): LinkSpec => ({ spid: String(n++), action })
    for (const verb of ['nextslide', 'previousslide', 'firstslide', 'lastslide', 'endshow'] as const) runLink(link({ kind: 'jump', verb }), nav, window, { sink, open })
    runLink(link({ kind: 'slide', index: 4 }), nav, window, { sink, open })
    runLink(link({ kind: 'url', href: 'mailto:a@example.com' }), nav, window, { sink, open })
    runLink(link({ kind: 'url', href: 'file:///etc/passwd' }), nav, window, { sink, open })
    runLink(link({ kind: 'refused', reason: 'ppaction://macro?name=x' }), nav, window, { sink, open })
    expect(calls).toEqual(['next', 'previous', 'first', 'last', 'exit', 'goTo:4'])
    expect(open.mock.calls).toEqual([[window, 'mailto:a@example.com', 'noopener']])
    expect(sink.list().filter(d => d.code === 'link-refused')).toHaveLength(2)
  })

  it('kiosk: links, triggers and their buttons are inert; the host API still navigates', async () => {
    const doc = await deck({
      presProps: '<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:showPr><p:kiosk/></p:showPr></p:presentationPr>',
      slides: [{
        shapes: [at(2, 400, 300, 200, 100, 'Revealed'), at(3, 0, 0, 96, 96, 'Press'), at(4, 200, 0, 96, 96, 'Web')],
        timingXml: timingXmlOf(interactive(20, 3, effect(22, 2, 'entr'))),
        links: [{ shapeId: 4, target: 'https://example.com/', external: true }],
      }, { shapes: [] }],
    })
    const r = await show(doc)
    expect(r.buttons()).toEqual([])
    await r.click(48, 48)
    await r.click(248, 48)
    expect(r.opened).toEqual([])
    expect(visible(stateNow(r), 'sp:2')).toBe(false)
    r.s.activate({ spid: '3', kind: 'trigger' })
    expect(visible(stateNow(r), 'sp:2')).toBe(false)
    // spec 4.4 names user input only: the host API is not gated
    r.s.next()
    await flush()
    expect(r.s.slideIndex).toBe(1)
    r.s.exit()
  })
})

describe('advance and timers', () => {
  it('advClick=0: click does not advance, Space does', async () => {
    const doc = await deck({ slides: [{ shapes: [at(2, 0, 0, 96, 96)], transition: { xml: '<p:transition advClick="0"><p:cut/></p:transition>' } }, { shapes: [] }] })
    const r = await show(doc)
    await r.click(500, 500)
    expect(r.s.slideIndex).toBe(0)
    await r.key(' ')
    expect(r.s.slideIndex).toBe(1)
    // the arriving slide has no advClick: clicks advance it (to the end of the show)
    await r.click(500, 500)
    expect(r.s.state.slide).toBe('end')
    r.s.exit()
  })

  it('auto-advance (advTm) and the kiosk restart run on the show clock and stop while paused', async () => {
    const doc = await deck({
      presProps: '<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:showPr><p:kiosk restart="5000"/></p:showPr></p:presentationPr>',
      slides: [
        { shapes: [], transition: { xml: '<p:transition advTm="1000"><p:cut/></p:transition>' } },
        { shapes: [] },
        { shapes: [] },
      ],
    })
    const r = await show(doc)
    await r.tick(999)
    expect(r.s.slideIndex).toBe(0)
    await r.tick(1)
    expect(r.s.slideIndex).toBe(1)
    // slide 1 has no advTm: it stays until the kiosk restart, 5000 ms after the start
    await r.tick(3000)
    expect(r.s.slideIndex).toBe(1)
    // pause stops the restart timer (Esc is the only kiosk key; pause through the clock as a hidden document would)
    clock.pause()
    await r.tick(10_000)
    expect(r.s.slideIndex).toBe(1)
    clock.resume()
    await r.tick(1000)
    expect(r.s.slideIndex).toBe(0)
    r.s.exit()
  })
})

describe('accessibility', () => {
  it('live region announces Slide n of m: title once per change', async () => {
    const doc = await deck({
      slides: [
        { shapes: [{ placeholder: { type: 'title' }, off: ['0', '0'], ext: ['9144000', '914400'], paragraphs: [{ runs: [{ text: 'Intro' }] }] }] },
        { shapes: [{ placeholder: { type: 'title' }, off: ['0', '0'], ext: ['9144000', '914400'], paragraphs: [{ runs: [{ text: 'Second' }] }] }] },
      ],
    })
    const r = await show(doc)
    const live = r.host.querySelector('[aria-live="polite"]')!
    expect(r.host.querySelectorAll('[aria-live]')).toHaveLength(1)
    expect(live.textContent).toBe('Slide 1 of 2: Intro')
    const records: string[] = []
    const mo = new MutationObserver(() => records.push(live.textContent ?? ''))
    mo.observe(live, { childList: true, characterData: true, subtree: true })
    for (let i = 0; i < 5; i++) await r.tick(16)
    await r.key('ArrowRight')
    for (let i = 0; i < 5; i++) await r.tick(16)
    await flush()
    expect(records).toEqual(['Slide 2 of 2: Second'])
    await r.key('ArrowRight')
    await flush()
    expect(live.textContent).toBe('End of slide show, click to exit.')
    mo.disconnect()
    r.s.exit()
  })

  it('offscreen text excludes not-yet-entered and exited text', async () => {
    const doc = await deck({
      slides: [{
        shapes: [at(2, 0, 0, 200, 50, 'Always'), at(4, 0, 100, 200, 50, 'Later'), at(5, 0, 200, 200, 50, 'Gone')],
        timingXml: timingXmlOf(mainSeqXml([step(10, effect(12, 4, 'entr') + effect(14, 5, 'exit'))])),
      }],
    })
    const r = await show(doc)
    expect(r.text()).toEqual(['Always', 'Gone'])
    await r.key('ArrowRight')
    await r.tick(10)
    expect(r.text()).toEqual(['Always', 'Later'])
    // Previous snaps back to the checkpoint: Later is hidden again, Gone is back
    await r.key('ArrowLeft')
    expect(r.text()).toEqual(['Always', 'Gone'])
    r.s.exit()
  })

  it('Tab reaches trigger and link buttons in z-order; Enter activates', async () => {
    const doc = await deck({
      slides: [{
        // paint order: the link (shape 6) below the trigger (shape 3)
        shapes: [at(6, 0, 300, 96, 96, 'Docs'), at(2, 400, 300, 200, 100, 'Revealed'), at(3, 0, 0, 96, 96, 'Press')],
        timingXml: timingXmlOf(interactive(20, 3, effect(22, 2, 'entr'))),
        links: [{ shapeId: 6, target: 'https://example.com/docs', external: true }],
      }, { shapes: [] }],
    })
    const r = await show(doc)
    const [link, trigger] = r.buttons()
    expect([link.textContent, trigger.textContent]).toEqual(['Link: Docs', 'Trigger: Shape 3'])
    for (const b of r.buttons()) { expect(b.tabIndex).toBe(0); expect(b.disabled).toBe(false) }
    // Tab is never taken by the show (focus is not trapped)
    const tab = new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true })
    link.dispatchEvent(tab)
    expect(tab.defaultPrevented).toBe(false)
    trigger.focus()
    await r.key('Enter', trigger)
    await r.tick(10)
    expect(visible(stateNow(r), 'sp:2')).toBe(true)
    expect(r.s.slideIndex).toBe(0)
    expect(document.activeElement?.textContent).toBe('Trigger: Shape 3')
    await r.key(' ', r.buttons()[0])
    expect(r.opened).toEqual([['https://example.com/docs', 'noopener']])
    expect(r.s.slideIndex).toBe(0)
    r.s.exit()
  })

  it('cursor hides after 3000 ms of no movement and returns on move', async () => {
    const r = await show(await deck({ slides: [{ shapes: [] }] }))
    const surface = r.surface()
    advanceWall(CURSOR_HIDE_MS - 1)
    expect(surface.hasAttribute('data-officeview-show-cursor')).toBe(false)
    advanceWall(1)
    expect(surface.getAttribute('data-officeview-show-cursor')).toBe('hidden')
    surface.dispatchEvent(new MouseEvent('mousemove', { bubbles: true }))
    expect(surface.hasAttribute('data-officeview-show-cursor')).toBe(false)
    advanceWall(CURSOR_HIDE_MS)
    expect(surface.getAttribute('data-officeview-show-cursor')).toBe('hidden')
    // the show clock does not drive it: pausing the show changes nothing
    expect(clock.now()).toBe(0)
    r.s.exit()
    expect(timers.filter(t => t.at > wall)).toHaveLength(0)
  })
})

describe('reduced motion', () => {
  const pushDeck = () => deck({
    slides: [
      { shapes: [at(2, 0, 0, 200, 100, 'One')] },
      {
        shapes: [at(2, 0, 0, 200, 100, 'Two'), at(3, 0, 200, 200, 100, 'Fades')],
        transition: { xml: '<p:transition spd="slow"><p:push dir="u"/></p:transition>' },
        timingXml: timingXmlOf(mainSeqXml([step(10, effect(12, 3, 'entr', fade(14, 3)))])),
      },
      { shapes: [] },
    ],
  })
  // jsdom has no matchMedia: the OS asks for reduced motion.
  const reducedQuery = () => Object.defineProperty(window, 'matchMedia', {
    configurable: true, writable: true,
    value: (q: string) => ({
      matches: q.includes('prefers-reduced-motion'), media: q, onchange: null,
      addEventListener() {}, removeEventListener() {}, addListener() {}, removeListener() {}, dispatchEvent: () => false,
    }) as MediaQueryList,
  })
  afterEach(() => { delete (window as { matchMedia?: unknown }).matchMedia })

  it('reduced motion: transitions crossfade within 200 ms, steps jump to end states', async () => {
    const r = await show(await pushDeck(), { motion: 'reduced' })
    await r.key('ArrowRight')
    expect(r.s.stage!.transition).toEqual({ effect: 'fade', duration: REDUCED_TRANSITION_MS })
    await r.tick(REDUCED_TRANSITION_MS)
    expect(r.s.stage!.transition).toBeUndefined()
    // the 1000 ms fade-in shows at its end state at once
    await r.key('ArrowRight')
    const s = stateNow(r).get('sp:3')!
    expect([s.visible, s.opacity]).toEqual([true, 1])
    // a cut stays a cut (reduced motion never adds motion)
    await r.key('ArrowRight')
    await r.key('ArrowRight')
    expect(r.s.slideIndex).toBe(2)
    expect(r.s.stage!.transition).toBeUndefined()
    r.s.exit()
  })

  it('slideShowMotion full overrides prefers-reduced-motion', async () => {
    reducedQuery()
    const full = await show(await pushDeck(), { motion: 'full' })
    await full.key('ArrowRight')
    expect(full.s.stage!.transition).toEqual({ effect: 'push', duration: 1000 })
    // mid-step the fade is still running under 'full'
    await full.tick(1000)
    await full.key('ArrowRight')
    await full.tick(500)
    expect(stateNow(full).get('sp:3')!.opacity).toBeLessThan(1)
    full.s.exit()
    // 'auto' follows the media query
    const auto = await show(await pushDeck(), { motion: 'auto' })
    await auto.key('ArrowRight')
    expect(auto.s.stage!.transition).toEqual({ effect: 'fade', duration: REDUCED_TRANSITION_MS })
    auto.s.exit()
  })
})

describe('output, residency and protection', () => {
  it('resize re-rasterises neighbours; old layers paint scaled until ready', async () => {
    // Twelve animated shapes per slide: more layers than the 8 re-rasterisations a frame allows.
    const shapes = Array.from({ length: 12 }, (_, i) => at(i + 2, (i % 4) * 220, Math.floor(i / 4) * 220, 200, 200, `S${i}`))
    const timing = timingXmlOf(mainSeqXml(shapes.map((_, i) => step(100 + i * 10, effect(102 + i * 10, i + 2, 'entr')))))
    const doc = await deck({ slides: [0, 1, 2].map(() => ({ shapes, timingXml: timing })) })
    const r = await show(doc, { from: 1 })
    await flush()
    await r.tick(16)
    const before = r.s.stage!.inspect()
    expect(before.slides.map(s => s.index)).toEqual([0, 1, 2])
    expect(before.slides.every(s => s.level === 'full' && s.pending === 0 && s.uploaded === s.layers)).toBe(true)
    const layers = before.slides[1].layers
    expect(layers).toBeGreaterThan(8)
    const drawn = before.lastDraws.length
    // a bigger window: a new reserve, then the output swaps
    vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(1920)
    vi.spyOn(window, 'innerHeight', 'get').mockReturnValue(1440)
    window.dispatchEvent(new Event('resize'))
    await flush()
    expect(r.s.stage!.backingStore).toMatchObject({ w: 1920, h: 1440 })
    expect(r.host.querySelector<HTMLCanvasElement>('[data-officeview-show-canvas]')!.width).toBe(1920)
    const pending = r.s.stage!.inspect()
    expect(pending.slides.every(s => s.pending === s.layers)).toBe(true)
    // one frame: at most 8 layers re-rasterised; every layer still draws (old ones scaled)
    r.s.tick()
    const one = r.s.stage!.inspect()
    expect(one.slides[1].pending).toBe(layers - 8)
    expect(one.lastDraws.length).toBe(drawn)
    for (let i = 0; i < 10; i++) r.s.tick()
    const done = r.s.stage!.inspect()
    expect(done.slides.every(s => s.pending === 0 && s.uploaded === s.layers)).toBe(true)
    expect(done.slides[1].scale).toBeCloseTo(2, 5)
    expect(done.slides[0].scale).toBeCloseTo(2, 5)
    expect(done.lastDraws.length).toBe(drawn)
    r.s.exit()
  })

  it('a resize the workspace cannot hold keeps the old backing store, scaled, with a diagnostic', async () => {
    vi.spyOn(sessionDeps, 'workspace').mockImplementation(() => createSlideShowWorkspace('desktop', { limit: 1024 * 768 * 4 * 2 + 64 * 2 ** 20 + 40 * 2 ** 20 }))
    const r = await show(await deck({ slides: [{ shapes: [at(2, 0, 0, 200, 200, 'A')] }] }))
    vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(3840)
    vi.spyOn(window, 'innerHeight', 'get').mockReturnValue(2160)
    window.dispatchEvent(new Event('resize'))
    await flush()
    expect(r.log).toContain('diag:resize-kept-old')
    expect(r.s.stage!.backingStore).toMatchObject({ w: 1024, h: 768 })
    r.s.exit()
  })

  it('watermark appears on the audience surface above animated layers', async () => {
    const doc = await deck({
      slides: [{ shapes: [at(2, 0, 0, 960, 720, 'Body')], timingXml: timingXmlOf(mainSeqXml([step(10, effect(12, 2, 'entr'))])) }],
    })
    const r = await show(doc, { watermark: { text: 'CONFIDENTIAL' } })
    await r.key('ArrowRight')
    await r.tick(10)
    const draws = r.s.stage!.inspect().lastDraws
    expect(draws[draws.length - 1]).toBe('watermark')
    expect(draws).toContain('target:sp:2')
    expect(draws.indexOf('target:sp:2')).toBeLessThan(draws.length - 1)
    // without a watermark there is no watermark layer
    r.s.exit()
    const plain = await show(doc)
    expect(plain.s.stage!.inspect().lastDraws).not.toContain('watermark')
    plain.s.exit()
  })

  it('allowCopy=false keeps the offscreen text readable but not selectable; allowPrint=false keeps slideshow surfaces out of print', async () => {
    // The viewer's theme root as OfficeDoc renders it with allowCopy={false} allowPrint={false}.
    const root = document.createElement('div')
    root.setAttribute('data-ov-theme-root', '')
    root.setAttribute('data-ov-print', 'false')
    const region = document.createElement('div')
    region.setAttribute('data-officeview-root', '')
    region.style.userSelect = 'none'
    root.append(region)
    document.body.append(root)
    const doc = await deck({ slides: [{ shapes: [at(2, 0, 0, 200, 50, 'Secret')] }] })
    const r = await show(doc, { host: hostIn(root) })
    const list = r.host.querySelector<HTMLElement>('[data-officeview-show-text]')!
    const item = list.querySelector('li')!
    expect(item.getAttribute('aria-label')).toBe('Secret')
    expect(item.textContent).toBe('')
    expect(list.style.userSelect).toBe('none')
    expect(r.surface().getAttribute('data-ov-print')).toBe('false')
    // the shipped stylesheet's print rule matches this surface and hides it
    const rules = Array.from(document.styleSheets).flatMap(sheet => Array.from(sheet.cssRules)).flatMap(rule =>
      rule instanceof CSSMediaRule && rule.media.mediaText.includes('print') ? Array.from(rule.cssRules) : [])
    const rule = rules.find(rule => rule instanceof CSSStyleRule && rule.selectorText.includes('data-officeview-show-surface')) as CSSStyleRule
    expect(rule.style.display).toBe('none')
    expect(rule.style.getPropertyPriority('display')).toBe('important')
    expect(r.surface().matches(rule.selectorText)).toBe(true)
    expect(readFileSync('src/styles/officeview.css', 'utf8')).toContain('[data-officeview-show-surface][data-ov-print="false"]')
    r.s.exit()
    // allowed: real text, selectable, printable
    const open = await show(doc)
    const li = open.host.querySelector('[data-officeview-show-text] li')!
    expect(li.textContent).toBe('Secret')
    expect(open.surface().hasAttribute('data-ov-print')).toBe(false)
    open.s.exit()
  })

  it('a tight workspace reaches the terminal rung: direct paint, cuts, no layers', async () => {
    // Room for the permanent reserve and the slides' metadata, not for one full-slide layer.
    vi.spyOn(sessionDeps, 'workspace').mockImplementation(() => createSlideShowWorkspace('desktop', { limit: 1024 * 768 * 4 + 8 * 2 ** 20 + 2 ** 20, scratch: 8 * 2 ** 20 }))
    const doc = await deck({
      slides: [{ shapes: [at(2, 0, 0, 480, 360, 'A')], timingXml: timingXmlOf(mainSeqXml([step(10, effect(12, 2, 'entr'))])) },
        { shapes: [at(2, 0, 0, 960, 720)], transition: { xml: '<p:transition><p:fade/></p:transition>' } }],
    })
    const r = await show(doc)
    expect(r.s.stage!.level).toBe('terminal')
    expect(r.log).toContain('diag:ladder-terminal')
    const ctx = r.host.querySelector<HTMLCanvasElement>('[data-officeview-show-canvas]')!.getContext('2d')!
    const px = (x: number, y: number) => [...ctx.getImageData(Math.round(x * K), Math.round(y * K), 1, 1).data]
    // the entrance has not run: the shape is not painted (no premature disclosure)
    expect(px(240, 180)).toEqual([255, 255, 255, 255])
    await r.key('ArrowRight')
    await r.tick(10)
    expect(px(240, 180)).toEqual([0x33, 0x66, 0x99, 255])
    await r.key('ArrowRight')
    expect(r.s.stage!.transition).toBeUndefined()
    expect(r.s.stage!.inspect().slides.every(s => s.uploaded === 0)).toBe(true)
    r.s.exit()
  })
})

describe('S1-19 carry-overs', () => {
  it('the post-exit focus restore listener goes after its timeout', async () => {
    const add = vi.spyOn(document, 'addEventListener')
    const remove = vi.fn()
    const r = await show(await deck({ slides: [{ shapes: [] }] }), { fullscreen: new Promise(() => {}) })
    r.s.exit()
    // the last one: the session's own full-screen resize listener ended with the show
    const listener = [...add.mock.calls].reverse().find(([type, , o]) => type === 'fullscreenchange' && (o as AddEventListenerOptions | undefined)?.signal)
    const signal = (listener![2] as AddEventListenerOptions).signal!
    signal.addEventListener('abort', remove)
    advanceWall(RESTORE_LISTENER_MS - 1)
    expect(signal.aborted).toBe(false)
    advanceWall(1)
    expect(signal.aborted).toBe(true)
  })

  it('layout and master artwork are parsed once per layout', async () => {
    const parse = await import('../src/pptx/parse')
    const theme = vi.spyOn(parse, 'slideTheme')
    const doc = await deck({ slides: [{ shapes: [] }, { shapes: [] }, { shapes: [] }] })
    const r = await show(doc)
    await r.key('ArrowRight')
    await r.key('ArrowRight')
    await flush()
    expect(theme.mock.calls.length).toBeLessThanOrEqual(1)
    r.s.exit()
  })

  it('a failed start still rejects typed and leaves the host empty', async () => {
    const doc = await deck({ slides: [{ shapes: [] }] })
    const host = hostIn()
    vi.spyOn(sessionDeps, 'stage').mockImplementation(() => { throw new SlideShowError('load-failed', 'boom') })
    await expect(show(doc, { host })).rejects.toMatchObject({ code: 'load-failed' })
    expect(host.childElementCount).toBe(0)
  })
})
