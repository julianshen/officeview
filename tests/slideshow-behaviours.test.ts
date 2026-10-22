// S1-11: behaviour evaluation to render state. Every oracle is numeric at a
// chosen time; the running engine consumes S1-9's snapshots directly.
import { describe, expect, it } from 'vitest'
import { parseXmlOrdered } from '../src/core/xml'
import { OfficePackage } from '../src/core/zip'
import { parsePptx } from '../src/pptx/parse'
import { buildPptxDeck, timingXmlOf } from '../src/testdata/ooxml-builders'
import { openShowPackage } from '../src/slideshow/read/deck'
import { readSlideTiming } from '../src/slideshow/read/timing'
import { shapeResolver } from '../src/slideshow/read/resolve'
import { parseMotionPath } from '../src/slideshow/motion-path'
import { createScheduler, type SchedulerSnapshot } from '../src/slideshow/engine/scheduler'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { targetKey } from '../src/slideshow/target'
import type { DiagnosticSink, Rng } from '../src/slideshow/types'
import type {
  BehaviourNode, CTn, ContainerNode, MotionPath, Point, ShowEvent,
  TargetRef, TargetState, TimeNode, TimingTree,
} from '../src/slideshow/types'
import { renderStateAt, rebuildFromPreset, standInFor, transitionStandIn, type BaseGeometry } from '../src/slideshow/engine/behaviours'

// ---------------------------------------------------------------------------
// Builders (the shapes the scheduler test uses)
// ---------------------------------------------------------------------------

type CtnOver = Partial<Omit<CTn, 'id'>>

function ctn(id: number, over: CtnOver = {}): CTn {
  return { id, durMs: undefined, accel: 0, decel: 0, autoRev: false, start: [], end: [], ...over }
}

function par(id: number, children: TimeNode[], over: CtnOver = {}): ContainerNode {
  return { kind: 'par', ctn: ctn(id, over), children }
}

function tree(...children: TimeNode[]): TimingTree {
  return { root: par(1, children, { durMs: 'indefinite', restart: 'never', nodeType: 'tmRoot' }), build: [] }
}

const T2: TargetRef = { spid: '2' }

/** A behaviour with the `p:cBhvr` defaults; `extra` covers the kind fields. */
function bhv(id: number, kind: BehaviourNode['kind'], extra: Record<string, unknown> = {}, ctnOver: CtnOver = {}): BehaviourNode {
  const base = { additive: 'base', accumulate: 'none' } as Pick<BehaviourNode, 'additive' | 'accumulate'>
  switch (kind) {
    case 'set':
      return { ...base, kind: 'set', ctn: ctn(id, ctnOver), target: T2, attrNames: ['style.visibility'], additive: 'base', accumulate: 'none', to: 'visible', ...extra } as unknown as BehaviourNode
    case 'anim':
      return { ...base, kind: 'anim', ctn: ctn(id, ctnOver), target: T2, attrNames: ['ppt_x'], additive: 'base', accumulate: 'none', calcmode: 'lin', keys: [], ...extra } as unknown as BehaviourNode
    case 'animRot':
      return { ...base, kind: 'animRot', ctn: ctn(id, ctnOver), target: T2, attrNames: ['r'], additive: 'base', accumulate: 'none', ...extra } as unknown as BehaviourNode
    case 'animScale':
      return { ...base, kind: 'animScale', ctn: ctn(id, ctnOver), target: T2, attrNames: ['ppt_width'], additive: 'base', accumulate: 'none', ...extra } as unknown as BehaviourNode
    case 'animMotion':
      return { ...base, kind: 'animMotion', ctn: ctn(id, ctnOver), target: T2, attrNames: ['ppt_x'], additive: 'base', accumulate: 'none', ...extra } as unknown as BehaviourNode
    case 'animEffect':
      return { ...base, kind: 'animEffect', ctn: ctn(id, ctnOver), target: T2, attrNames: ['style.visibility'], additive: 'base', accumulate: 'none', transition: 'in', ...extra } as unknown as BehaviourNode
    case 'animClr':
      return { ...base, kind: 'animClr', ctn: ctn(id, ctnOver), target: T2, attrNames: ['fillcolor'], additive: 'base', accumulate: 'none', clrSpc: 'rgb', dir: 'cw', ...extra } as unknown as BehaviourNode
    default:
      throw new Error(`no builder for ${kind}`)
  }
}

const k = (tm: number, value: string, fmla?: string): { tm: number; value: string; fmla?: string } => ({ tm, value, ...(fmla === undefined ? {} : { fmla }) })

const SLIDE = { w: 960, h: 720 }
const rng: Rng = () => 0.5
const geo = (over: Partial<BaseGeometry> = {}): BaseGeometry => ({ x: 0, y: 108, w: 48, h: 36, rot: 0, ...over })
const bases = (g: BaseGeometry, ref: TargetRef = T2) => new Map([[targetKey(ref), g]])

/** Runs a hand-built tree to time `at` and renders the state of T2. */
function run(children: TimeNode[], at: number, opts: {
  events?: readonly ShowEvent[]
  basis?: BaseGeometry
  slide?: { w: number; h: number }
} = {}): { snap: SchedulerSnapshot; base: TargetState | undefined; sink: DiagnosticSink } {
  const sink = createDiagnosticSink()
  const sched = createScheduler(tree(...children))
  const snap = sched.evaluate(opts.events ?? [], at)
  const map = renderStateAt(snap, bases(opts.basis ?? geo()), rng, sink, opts.slide ?? SLIDE)
  return { snap, base: map.get(targetKey(T2)), sink }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('behaviour evaluation', () => {
  it('anim ppt_x from #ppt_x-1 to #ppt_x at 0, 0.5 and 1', () => {
    const fly = bhv(11, 'anim', {
      target: T2, attrNames: ['ppt_x'], calcmode: 'lin',
      keys: [k(0, '#ppt_x-1'), k(1, '#ppt_x')],
    }, { durMs: 1000, fill: 'hold' })
    const step = par(10, [fly], { durMs: 1000, fill: 'hold' })
    expect(run([step], 0).base?.motion?.dx).toBeCloseTo(-960)
    expect(run([step], 500).base?.motion?.dx).toBeCloseTo(-480)
    expect(run([step], 1000).base?.motion?.dx).toBeCloseTo(0)
    // #ppt_x inside the formula reads the target's own base: a shape at x = 192 px starts 1 width left of THERE.
    const shifted = run([step], 0, { basis: geo({ x: 192 }), slide: SLIDE }).base
    expect(shifted?.motion?.dx).toBeCloseTo(-960)
    expect(shifted?.motion?.dy ?? 0).toBeCloseTo(0)
  })

  it('tav keyframes with fmla and calcmode lin and discrete', () => {
    // Linear: tav segments interpolate; a key carrying `fmla` replaces its value at its own tm with the formula result.
    const lin = bhv(11, 'anim', {
      target: T2, attrNames: ['style.opacity'], calcmode: 'lin',
      keys: [k(0, '0'), { tm: 0.5, value: '0', fmla: '$+0.2' }, k(1, '1')],
    }, { durMs: 1000, fill: 'hold' })
    const linWrap = par(20, [lin], { durMs: 1000, fill: 'hold' })
    expect(run([linWrap], 500).base?.opacity).toBeCloseTo(0.7)
    expect(run([linWrap], 250).base?.opacity).toBeCloseTo(0.35)
    expect(run([linWrap], 750).base?.opacity).toBeCloseTo(0.85)
    expect(run([linWrap], 1000).base?.opacity).toBeCloseTo(1)
    // Discrete: holds the value of the last key whose tm <= the posed progress.
    const disc = bhv(11, 'anim', {
      target: T2, attrNames: ['style.opacity'], calcmode: 'discrete',
      keys: [k(0, '0'), k(0.5, '0.4'), k(1, '0.9')],
    }, { durMs: 1000, fill: 'hold' })
    const discWrap = par(20, [disc], { durMs: 1000, fill: 'hold' })
    expect(run([discWrap], 250).base?.opacity).toBeCloseTo(0)
    expect(run([discWrap], 750).base?.opacity).toBeCloseTo(0.4)
    expect(run([discWrap], 1000).base?.opacity).toBeCloseTo(0.9)
  })

  it('animRot by 90 at half time is 45; animScale 150% with zoomContents', () => {
    // Rotation composes additively onto the base rotation: base 20° is at 65° halfway (cos < 0.5).
    const spin = bhv(11, 'animRot', { by: 90 }, { durMs: 1000, fill: 'hold' })
    const spinWrap = par(20, [spin], { durMs: 1000, fill: 'hold' })
    expect(run([spinWrap], 500).base?.transform[0]).toBeCloseTo(0.7071067812) // cos45
    expect(run([spinWrap], 500).base?.transform[1]).toBeCloseTo(0.7071067812) // sin45
    expect(run([spinWrap], 1000).base?.transform[0]).toBeCloseTo(0) // cos90
    const spun = run([spinWrap], 500, { basis: geo({ rot: 20 }) }).base
    expect(spun?.transform[0]).toBeLessThan(0.5)
    expect(spun?.transform[1]).toBeGreaterThan(0.9)
    // Scale about the centre: at half time the factor is 1.25; tx = cx - s*cx.
    const zoom = bhv(11, 'animScale', { by: { x: 1.5, y: 1.5 } as Point, zoomContents: true }, { durMs: 1000, fill: 'hold' })
    const zoomWrap = par(20, [zoom], { durMs: 1000, fill: 'hold' })
    const scaled = run([zoomWrap], 500).base
    expect(scaled?.transform[0]).toBeCloseTo(1.25)
    expect(scaled?.transform[3]).toBeCloseTo(1.25)
    expect(scaled?.transform[4]).toBeCloseTo(24 - 1.25 * 24) // cx = 48/2
    expect(scaled?.transform[5]).toBeCloseTo((108 + 18) - 1.25 * (108 + 18)) // cy = y + h/2
    expect(run([zoomWrap], 0).base?.transform[0]).toBeCloseTo(1)
  })

  it('animMotion follows the path and rAng; origin layout vs parent', () => {
    const parsed = parseMotionPath('M 0 0 L 1 0')
    const path: MotionPath = parsed.ok ? parsed.path : { segments: [], length: 0 }
    const base = geo({ x: 100, y: 100 })
    const layout = bhv(11, 'animMotion', { path, origin: 'layout' }, { durMs: 1000, fill: 'hold' })
    const parent = bhv(11, 'animMotion', { path, origin: 'parent' }, { durMs: 1000, fill: 'hold' })
    const turned = bhv(11, 'animMotion', { path, origin: 'layout', rAngDeg: 90 }, { durMs: 1000, fill: 'hold' })
    const P = (b: BehaviourNode, t: number) => run([par(20, [b], { durMs: 1000 })], t, { basis: base }).base?.motion
    // origin layout: path coordinates are absolute slide fractions; the offset is P*slide - base.
    expect(P(layout, 250)?.dx).toBeCloseTo(0.25 * SLIDE.w - 100)
    expect(P(layout, 250)?.dy).toBeCloseTo(-100)
    // origin parent: offsets already measured from the shape's own position.
    expect(P(parent, 250)?.dx).toBeCloseTo(0.25 * SLIDE.w)
    expect(P(parent, 250)?.dy).toBeCloseTo(0)
    // rAng 90° rotates the offset about the origin: (140, -100) -> (100, 140).
    expect(P(turned, 250)?.dx).toBeCloseTo(100)
    expect(P(turned, 250)?.dy).toBeCloseTo(140)
  })

  it('animEffect wipe(down) progress maps to the mask', () => {
    const wipe = bhv(11, 'animEffect', { filter: 'wipe(down)', transition: 'in' }, { durMs: 1000, fill: 'hold' })
    const wipeWrap = par(20, [wipe], { durMs: 1000, fill: 'hold' })
    const mid = run([wipeWrap], 500).base
    expect(mid?.mask?.kind).toBe('wipe')
    expect(mid?.mask?.progress).toBeCloseTo(0.5)
    expect(mid?.mask?.params.from).toBe('t')
    expect(run([wipeWrap], 1000).base?.mask?.progress).toBeCloseTo(1)
    // A static progress attr overrides the posed progress.
    const fixed = bhv(11, 'animEffect', { filter: 'wipe(down)', transition: 'in', progress: '0.4' }, { durMs: 1000, fill: 'hold' })
    expect(run([par(20, [fixed], { durMs: 1000 })], 800).base?.mask?.progress).toBeCloseTo(0.4)
    // The fade filter has no mask: it is an opacity ramp; 'out' ramps down.
    const fade = bhv(11, 'animEffect', { filter: 'fade()', transition: 'out' }, { durMs: 1000, fill: 'hold' })
    const f = run([par(20, [fade], { durMs: 1000 })], 500).base
    expect(f?.mask).toBeUndefined()
    expect(f?.opacity).toBeCloseTo(0.5)
    // An unknown filter is diagnosed, never silently nothing (renders as a fade).
    const unknown = bhv(11, 'animEffect', { filter: 'melt()' }, { durMs: 1000, fill: 'hold' })
    const u = run([par(20, [unknown], { durMs: 1000 })], 500)
    expect(u.base?.mask).toBeUndefined()
    expect(u.base?.opacity).toBeCloseTo(0.5)
    expect(u.sink.list().some(d => d.code === 'anim-effect-unknown-filter')).toBe(true)
  })

  it('animClr hsl dir cw and rgb by; fill-only, line-only and text-only addressing', () => {
    // rgb `by` offsets scale the r channel: fill[5] = 1 + by.r*p.
    const scaleBy = bhv(11, 'animClr', { clrSpc: 'rgb', dir: 'cw', by: 'rgb(0.25,0,0)', attrNames: ['fillcolor'] }, { durMs: 1000, fill: 'hold' })
    const scaleWrap = par(20, [scaleBy], { durMs: 1000, fill: 'hold' })
    const mid = run([scaleWrap], 500).base
    expect(mid?.color?.fill?.[0]).toBeCloseTo(1.125) // r gain, R output row
    expect(mid?.color?.fill?.[6]).toBeCloseTo(1)    // g -> g
    expect(mid?.color?.fill?.[7]).toBeCloseTo(0)    // g <- b
    expect(mid?.color?.line).toBeUndefined()
    expect(mid?.color?.text).toBeUndefined()
    // from->to slides the channel scales toward the destination colour's own channels.
    const fromTo = bhv(11, 'animClr', { clrSpc: 'rgb', dir: 'cw', from: '#ff0000', to: '#0000ff', attrNames: ['fillcolor'] }, { durMs: 1000, fill: 'hold' })
    const fromWrap = par(20, [fromTo], { durMs: 1000, fill: 'hold' })
    const steep = run([fromWrap], 500).base?.color?.fill
    expect(steep?.[0]).toBeCloseTo(0.5)   // r 1 -> 0
    expect(steep?.[6]).toBeCloseTo(1)     // g unchanged
    expect(steep?.[14]).toBeCloseTo(0.5)  // b offset 0 -> 0.5
    // The same behaviour addressed to stroke lands only on color.line.
    const stroke = bhv(11, 'animClr', { clrSpc: 'rgb', dir: 'cw', from: '#ff0000', to: '#0000ff', attrNames: ['stroke.color'] }, { durMs: 1000, fill: 'hold' })
    const strokeState = run([par(20, [stroke], { durMs: 1000 })], 500).base
    expect(strokeState?.color?.fill).toBeUndefined()
    expect(strokeState?.color?.line?.[0]).toBeCloseTo(0.5)
    const text = bhv(11, 'animClr', { clrSpc: 'rgb', dir: 'cw', from: '#ff0000', to: '#0000ff', attrNames: ['style.color'] }, { durMs: 1000, fill: 'hold' })
    const texState = run([par(20, [text], { durMs: 1000 })], 500).base
    expect(texState?.color?.text?.[0]).toBeCloseTo(0.5)
    expect(texState?.color?.fill).toBeUndefined()
  })

  it('additive sum and mult compose in document order', () => {
    const add = bhv(11, 'anim', { attrNames: ['style.opacity'], additive: 'sum', keys: [k(0, '0.25'), k(1, '0.25')] }, { durMs: 1000, fill: 'hold' })
    const mul = bhv(12, 'anim', { attrNames: ['style.opacity'], additive: 'mult', keys: [k(0, '0.5'), k(1, '0.5')] }, { durMs: 1000, fill: 'hold' })
    const wrap = par(20, [add, mul], { durMs: 1000, fill: 'hold' })
    // Document order: (1 + 0.25) * 0.5 = 0.625, not the reverse composition 0.75.
    expect(run([wrap], 0).base?.opacity).toBeCloseTo(0.625)
  })

  it('accumulate=always with repeat 3 holds the accumulated value after the end', () => {
    const { base: frozenAlways } = run([par(20, [bhv(11, 'anim', { target: T2, attrNames: ['ppt_x'], calcmode: 'lin', keys: [k(0, '0'), k(1, '0.25')], accumulate: 'always' }, { durMs: 1000, repeatCount: 3, fill: 'hold' })], { durMs: 3000, fill: 'hold' })], 3000)
    // Three iterations of +0.25 of one slide width each.
    expect(frozenAlways?.motion?.dx).toBeCloseTo(0.75 * SLIDE.w)
    // Without accumulation the frozen value is the single-iteration end.
    const { base: frozenOnce } = run([par(20, [bhv(11, 'anim', { target: T2, attrNames: ['ppt_x'], calcmode: 'lin', keys: [k(0, '0'), k(1, '0.25')] }, { durMs: 1000, repeatCount: 3, fill: 'hold' })], { durMs: 3000, fill: 'hold' })], 3000)
    expect(frozenOnce?.motion?.dx).toBeCloseTo(0.25 * SLIDE.w)
  })

  it('a preset with missing behaviours is rebuilt and diagnosed approximate', () => {
    const preset = ctn(10, { presetClass: 'entr', presetId: 22, presetSubtype: 8 })
    const sink = createDiagnosticSink()
    const last = rebuildFromPreset(preset, T2, sink)
    expect(last && last.length > 0).toBe(true)
    expect(sink.list().filter(d => d.code === 'preset-approximate')).toHaveLength(1)
    const again = rebuildFromPreset(preset, T2, sink)
    expect(again?.length).toBe(last?.length)
    expect(sink.list().filter(d => d.code === 'preset-approximate')).toHaveLength(1)
  })

  it('stand-ins: entrance visible, emphasis none, exit hidden, path end position; each diagnosed once', () => {
    expect(standInFor('entr')).toBe('visible')
    expect(standInFor('emph')).toBe('none')
    expect(standInFor('exit')).toBe('hidden')
    expect(standInFor('path')).toBe('path-end')
    expect(transitionStandIn()).toBe('fade')
  })

  it('an invalid formula drops only its behaviour', () => {
    const good = bhv(11, 'set', { to: 0.5, attrNames: ['style.opacity'] }, { durMs: 1000, fill: 'hold' })
    const bad = bhv(12, 'anim', {
      target: T2, attrNames: ['ppt_x'], calcmode: 'lin',
      keys: [k(0, '#ppt_x-1'), k(1, '#no_such_name')],
    }, { durMs: 1000, fill: 'hold' })
    // The bad behaviour alone: state is absent, diagnosed once.
    const dropped = run([par(20, [bad], { durMs: 1000 })], 500)
    expect(dropped.base).toBeUndefined()
    expect(dropped.sink.list().filter(d => d.code === 'behaviour-invalid-formula')).toHaveLength(1)
    // Re-running the same evaluation does not report it twice.
    expect(run([par(20, [bad], { durMs: 1000 })], 250).sink.list().filter(d => d.code === 'behaviour-invalid-formula')).toHaveLength(1)
    // The still-valid neighbour renders alongside it unchanged.
    const both = run([par(20, [good, bad], { durMs: 1000 })], 0)
    expect(both.base).toBeDefined()
    expect(both.base?.opacity).toBeCloseTo(0.5)
    expect(both.sink.list().filter(d => d.code === 'behaviour-invalid-formula')).toHaveLength(1)
  })

  it('a builder deck with an onMouseOver trigger goes reader to scheduler to render state', async () => {
    const deck = await buildPptxDeck({
      slides: [
        {
          shapes: [
            { id: 2, name: 'Target', off: ['0', '1149350'], ext: ['571500', '285750'] },
            { id: 3, name: 'Hover', off: ['5715000', '1149350'], ext: ['285750', '285750'] },
          ],
          timingXml: timingXmlOf(
            `<p:par><p:cTn id="10" fill="hold"><p:stCondLst><p:cond evt="onMouseOver" delay="0"><p:tgtEl><p:spTgt spid="3"/></p:tgtEl></p:cond></p:stCondLst>` +
            `<p:childTnLst><p:par><p:cTn id="11" presetID="2" presetClass="entr" presetSubtype="8" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst>` +
            `<p:childTnLst><p:anim><p:cBhvr><p:cTn id="12" dur="1000" fill="hold"/><p:tgtEl><p:spTgt spid="2"/></p:tgtEl>` +
            `<p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst>` +
            `<p:tav tm="0"><p:val><p:strVal val="#ppt_x-1"/></p:val></p:tav>` +
            `<p:tav tm="100000"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav></p:tavLst></p:anim></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`,
          ),
        },
      ],
    })
    const handle = await openShowPackage(deck)
    const pkg = await OfficePackage.load(deck)
    const doc = await parsePptx(pkg)
    const slideXml = await pkg.text(handle.slideParts[0])
    const node = parseXmlOrdered(slideXml!)
    const sink = createDiagnosticSink()
    const read = readSlideTiming(node, 0, shapeResolver(doc.slides[0]), sink)
    expect(sink.list().filter(d => d.code.startsWith('timing-'))).toHaveLength(0)
    expect(read.timing).toBeDefined()
    // The slide is 9144000x6858000 EMU = 960x720 px; the shape sits at x=0.
    const baseMap = new Map([[targetKey(T2), { x: 0, y: 120, w: 60, h: 30, rot: 0 } as BaseGeometry]])
    const timing = read.timing
    if (!timing || 'refused' in timing) throw new Error('timing was refused')
    const sched = createScheduler({ root: timing.root, build: [] })
    const events = [{ kind: 'mouseOver' as const, spid: '3', t: 250 }]
    const mid = renderStateAt(sched.evaluate(events, 1250), baseMap, rng, sink, SLIDE).get(targetKey(T2))
    // Fire at 250, evaluated at 1250: the end state sits at the shape's own x again.
    expect(mid?.motion?.dx).toBeCloseTo(0)
    expect(sink.list().filter(d => d.code === 'behaviour-invalid-formula')).toHaveLength(0)
  })
})
