/**
 * S1-14: the WebGL2 compositor's lifecycle and the backend manager's
 * context-loss handling (slideshow spec 6.3, 7). Node has no WebGL2, so the
 * GL is a recording fake: these tests pin order, leases and resource
 * recreation, never pixels. Pixels (and WebGL2 ↔ Canvas 2D parity) are
 * checked in Chromium by tests/browser/slideshow-gpu.spec.ts.
 *
 * The workspace is a test ledger typed by the S1-2 contracts (S1-16 owns the
 * real one), the same shape S1-12's and S1-13's tests use. Every clock is
 * injected: `now` and `schedule` are a manual queue, never wall time.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { ManualClock } from '../src/slideshow/clock'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { paintDirect, type DirectAssets } from '../src/slideshow/render/direct'
import { createWebGL2Compositor } from '../src/slideshow/render/webgl2'
import { RESTORE_DEADLINE_MS, createBackendManager, createLivePainter, type BackendManager, type LiveScene } from '../src/slideshow/render/backend'
import { frameDiff, withinTolerance } from '../src/debug/slideshow-parity'
import type { PptxSlide } from '../src/pptx/types'
import type { Compositor, DiagnosticSink, LayerHandle, Lease, SlideShowWorkspace, SubOwner, TerminalReserve } from '../src/slideshow/types'

// ---------------------------------------------------------------------------
// Workspace ledger
// ---------------------------------------------------------------------------
interface Pool { used: number; limit: number }
function poolLease(pool: Pool, bytes: number, tag: string, log?: string[]): Lease | undefined {
  if (bytes < 0 || pool.used + bytes > pool.limit) { log?.push(`refused:${tag}`); return undefined }
  pool.used += bytes
  log?.push(`lease:${tag}`)
  let held = bytes, live = true
  const lease: Lease = {
    get tag() { return tag },
    get bytes() { return live ? held : 0 },
    grow(extra) {
      if (!live || pool.used + extra > pool.limit) { log?.push(`grow-refused:${tag}`); return false }
      pool.used += extra; held += extra; log?.push(`grow:${tag}`); return true
    },
    shrink(less) { const n = Math.min(less, held); held -= n; pool.used -= n },
    transfer() { return lease },
    release() { if (!live) return; live = false; pool.used -= held; log?.push(`release:${tag}`) },
  }
  return lease
}
function makeWorkspace(limit: number, log?: string[]): SlideShowWorkspace & { pool: Pool; terminal?: TerminalReserve & { sub: Pool } } {
  const pool: Pool = { used: 0, limit }
  const ws: SlideShowWorkspace & { pool: Pool; terminal?: TerminalReserve & { sub: Pool } } = {
    pool, limit,
    lease: (bytes, tag) => poolLease(pool, bytes, tag, log),
    used: () => pool.used,
    reserveTerminal(outputBytes) {
      const scratchBytes = 1 << 20
      const whole = poolLease(pool, outputBytes + scratchBytes, 'terminal', log)
      if (!whole) return undefined
      const sub: Pool = { used: 0, limit: scratchBytes }
      const scratch: SubOwner = { lease: (bytes, tag) => poolLease(sub, bytes, `scratch:${tag}`, log) }
      ws.terminal = { outputBytes, scratch, sub, resize: () => false, release: () => whole.release() }
      return ws.terminal
    },
  }
  return ws
}

// ---------------------------------------------------------------------------
// A recording fake WebGL2 context
// ---------------------------------------------------------------------------
interface GlCall { name: string; args: unknown[] }
interface FakeGL {
  gl: WebGL2RenderingContext
  calls: GlCall[]
  count(name: string): number
  names(): string[]
  lose(): void
  restore(): void
  constant(name: string): number
}
function fakeGL(log?: string[], maxTextureSize = 4096): FakeGL {
  const calls: GlCall[] = []
  const constants = new Map<string, number>()
  let lost = false
  let serial = 0
  const constant = (name: string) => {
    let v = constants.get(name)
    if (v === undefined) { v = 0x1000 + constants.size; constants.set(name, v) }
    return v
  }
  const overrides: Record<string, unknown> = {
    isContextLost: () => lost,
    getParameter: (p: number) => (p === constant('MAX_TEXTURE_SIZE') ? maxTextureSize : 0),
    getShaderParameter: () => true,
    getProgramParameter: () => true,
    getShaderInfoLog: () => '',
    getProgramInfoLog: () => '',
    getUniformLocation: (_p: unknown, name: string) => ({ uniform: name }),
    getAttribLocation: () => 0,
    getExtension: () => null,
    getError: () => 0,
  }
  const gl = new Proxy({}, {
    get(_t, prop) {
      if (typeof prop !== 'string') return undefined
      if (prop in overrides) return overrides[prop]
      if (/^[A-Z][A-Z0-9_]*$/.test(prop)) return constant(prop)
      return (...args: unknown[]) => {
        calls.push({ name: prop, args })
        log?.push(`gl:${prop}`)
        if (prop.startsWith('create')) return { kind: prop.slice(6), serial: ++serial }
        return undefined
      }
    },
  }) as WebGL2RenderingContext
  return {
    gl, calls,
    count: name => calls.filter(c => c.name === name).length,
    names: () => calls.map(c => c.name),
    lose() { lost = true },
    restore() { lost = false },
    constant,
  }
}

let current: FakeGL | null = null
let realGetContext: typeof HTMLCanvasElement.prototype.getContext
beforeEach(() => {
  realGetContext = HTMLCanvasElement.prototype.getContext
  current = null
  // jsdom has no WebGL2: hand out the current fake for 'webgl2', node-canvas for '2d'.
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(function (this: HTMLCanvasElement, type: string, ...rest: unknown[]) {
    if (type === 'webgl2') return (current?.gl ?? null) as never
    return (realGetContext as (...a: unknown[]) => unknown).call(this, type, ...rest) as never
  } as never)
})
afterEach(() => { vi.restoreAllMocks(); document.body.replaceChildren() })

/** A solid layer charged to its own lease, then uploaded. */
function layerOn(c: Compositor, ws: SlideShowWorkspace, w = 8, h = 8): LayerHandle | { refused: string } {
  const canvas = document.createElement('canvas')
  canvas.width = w; canvas.height = h
  const x = canvas.getContext('2d')!
  x.fillStyle = '#ff0000'; x.fillRect(0, 0, w, h)
  const bytes = w * h * 4
  const lease = ws.lease(bytes, 'layer')
  if (!lease) return { refused: 'workspace' }
  const r = c.upload({ canvas, bounds: { x: 0, y: 0, w, h }, scale: 1, bytes }, lease)
  if ('refused' in r) lease.release()
  return r
}

// ---------------------------------------------------------------------------
// Manager harness
// ---------------------------------------------------------------------------
function harness(opts: { limit?: number; maxTextureSize?: number; gl?: boolean } = {}) {
  const log: string[] = []
  const ws = makeWorkspace(opts.limit ?? 256 * 1024 * 1024, log)
  const terminal = ws.reserveTerminal(320 * 180 * 4)!
  const fake = fakeGL(log, opts.maxTextureSize)
  current = opts.gl === false ? null : fake
  // As on the audience surface: the terminal output canvas (the permanent reserve's surface), then the host.
  const root = document.createElement('div')
  const output = document.createElement('canvas')
  output.width = 320; output.height = 180
  const host = document.createElement('div')
  root.append(output, host)
  document.body.appendChild(root)
  let t = 0
  const queue: Array<{ at: number; cb: () => void }> = []
  const sink: DiagnosticSink = createDiagnosticSink()
  const rebuilt: Array<Compositor['backend']> = []
  const rebuildGate: { hold?: Promise<void> } = {}
  const paintLive = vi.fn((ctx: CanvasRenderingContext2D) => { ctx.fillStyle = '#00ff00'; ctx.fillRect(0, 0, 4, 4) })
  const deps = {
    now: () => t,
    schedule: (cb: () => void, ms: number) => { queue.push({ at: t + ms, cb }); return queue.length },
    paintLive,
    rebuildLayers: vi.fn(async (c: Compositor) => {
      rebuilt.push(c.backend)
      if (rebuildGate.hold) await rebuildGate.hold
      const r = layerOn(c, ws)
      if ('refused' in r) throw new Error(`rebuild refused ${r.refused}`)
    }),
    workspace: ws,
    terminal,
    output,
    sink,
  }
  const mgr: BackendManager = createBackendManager(host, deps)
  /** Move the monotonic clock and run every scheduled callback now due. */
  const advance = (ms: number) => {
    t += ms
    for (;;) {
      const due = queue.filter(q => q.at <= t).sort((a, b) => a.at - b.at)[0]
      if (!due) break
      queue.splice(queue.indexOf(due), 1)
      due.cb()
    }
  }
  /** Fire every scheduled callback now, early, without moving the clock (a throttled or misbehaving timer). */
  const fireEarly = () => {
    const all = queue.splice(0)
    for (const q of all) q.cb()
  }
  const jump = (ms: number) => { t += ms }
  const glCanvas = () => host.querySelector<HTMLCanvasElement>('canvas[data-slideshow-backend="webgl2"]')
  const liveCanvas = () => root.querySelector<HTMLCanvasElement>('canvas[data-slideshow-backend="canvas2d"]')
  const lose = () => {
    const canvas = glCanvas()!
    fake.lose()
    const e = new Event('webglcontextlost', { cancelable: true })
    canvas.dispatchEvent(e)
    return e
  }
  const restore = (canvas = glCanvas()) => {
    fake.restore()
    canvas?.dispatchEvent(new Event('webglcontextrestored'))
  }
  const settle = async () => { for (let i = 0; i < 5; i++) await Promise.resolve() }
  return { log, ws, terminal, fake, host, output, deps, mgr, advance, fireEarly, jump, glCanvas, liveCanvas, lose, restore, settle, sink, rebuilt, rebuildGate, paintLive, now: () => t }
}

// ---------------------------------------------------------------------------
// The WebGL2 compositor
// ---------------------------------------------------------------------------
describe('createWebGL2Compositor', () => {
  it('returns undefined when there is no WebGL2 context', () => {
    current = null
    const canvas = document.createElement('canvas')
    expect(createWebGL2Compositor(canvas, makeWorkspace(1 << 30))).toBeUndefined()
  })

  it('texture upload leases before texImage2D and refuses without calling it', () => {
    const log: string[] = []
    const layerBytes = 8 * 8 * 4
    const fake = fakeGL(log)
    current = fake
    const canvas = document.createElement('canvas')
    canvas.width = 16; canvas.height = 16
    const ws = makeWorkspace(1 << 30, log)
    const comp = createWebGL2Compositor(canvas, ws)!
    expect(comp.backend).toBe('webgl2')
    expect(comp.maxTextureSize).toBe(4096)

    log.length = 0
    const ok = layerOn(comp, ws)
    expect('id' in ok).toBe(true)
    const grow = log.indexOf('grow:layer')
    expect(grow).toBeGreaterThanOrEqual(0)
    expect(log.indexOf('gl:createTexture')).toBeGreaterThan(grow)
    expect(log.indexOf('gl:texImage2D')).toBeGreaterThan(grow)
    // Premultiplied upload, set before the texture data goes in.
    const tex = fake.calls.findIndex(c => c.name === 'texImage2D')
    const premul = fake.calls.findIndex(c => c.name === 'pixelStorei' && c.args[0] === fake.constant('UNPACK_PREMULTIPLY_ALPHA_WEBGL') && c.args[1] === true)
    expect(premul).toBeGreaterThanOrEqual(0)
    expect(premul).toBeLessThan(tex)
    // The CPU copy's bytes leave the lease after upload; the texture's stay.
    const handle = ok as LayerHandle
    const usedWithTexture = ws.used()
    comp.release(handle)
    expect(usedWithTexture - ws.used()).toBe(layerBytes)

    // A workspace that holds the CPU copy but not the texture: refused, nothing created.
    const tight = makeWorkspace(0, log)
    const fake2 = fakeGL(log)
    current = fake2
    const comp2 = createWebGL2Compositor(document.createElement('canvas'), makeWorkspace(1 << 30))!
    const before = fake2.count('texImage2D')
    const createdBefore = fake2.count('createTexture')
    tight.pool.limit = layerBytes
    const lease = tight.lease(layerBytes, 'layer')!
    const c = document.createElement('canvas'); c.width = 8; c.height = 8
    expect(comp2.upload({ canvas: c, bounds: { x: 0, y: 0, w: 8, h: 8 }, scale: 1, bytes: layerBytes }, lease)).toEqual({ refused: 'workspace' })
    expect(fake2.count('texImage2D')).toBe(before)
    expect(fake2.count('createTexture')).toBe(createdBefore)
    expect(lease.bytes).toBe(layerBytes)

    // Over MAX_TEXTURE_SIZE: refused before any lease or GL call.
    const big = document.createElement('canvas'); big.width = 4097; big.height = 1
    const bigLease = makeWorkspace(1 << 30).lease(4097 * 4, 'big')!
    expect(comp2.upload({ canvas: big, bounds: { x: 0, y: 0, w: 4097, h: 1 }, scale: 1, bytes: 4097 * 4 }, bigLease)).toEqual({ refused: 'texture-size' })
    expect(bigLease.bytes).toBe(4097 * 4)
    expect(fake2.count('texImage2D')).toBe(before)
  })

  it('blends ONE, ONE_MINUS_SRC_ALPHA and draws one quad per visible layer', () => {
    const fake = fakeGL()
    current = fake
    const ws = makeWorkspace(1 << 30)
    const comp = createWebGL2Compositor(document.createElement('canvas'), ws)!
    comp.setViewport({ matrix: [2, 0, 0, 2, 0, 0], deviceWidth: 32, deviceHeight: 32 })
    const a = layerOn(comp, ws) as LayerHandle
    const b = layerOn(comp, ws) as LayerHandle
    fake.calls.length = 0
    const s = { visible: true, opacity: 1, transform: [1, 0, 0, 1, 0, 0] as const }
    comp.frame([{ handle: a, state: s }, { handle: b, state: { ...s, visible: false } }, { handle: b, state: { ...s, opacity: 0.5 } }])
    const blend = fake.calls.find(c => c.name === 'blendFunc')!
    expect(blend.args).toEqual([fake.constant('ONE'), fake.constant('ONE_MINUS_SRC_ALPHA')])
    expect(fake.count('drawArrays')).toBe(2)
    const vp = fake.calls.find(c => c.name === 'viewport')!
    expect(vp.args).toEqual([0, 0, 32, 32])
  })

  it('leases every mask surface before creating it, and a refused mask lease is a diagnosed cut', () => {
    const log: string[] = []
    const fake = fakeGL(log)
    current = fake
    const ws = makeWorkspace(1 << 30, log)
    const sink = createDiagnosticSink()
    const comp = createWebGL2Compositor(document.createElement('canvas'), ws, { sink })!
    const a = layerOn(comp, ws) as LayerHandle
    log.length = 0
    const s = { visible: true, opacity: 1, transform: [1, 0, 0, 1, 0, 0] as const }
    comp.frame([{ handle: a, state: { ...s, mask: { kind: 'circle', progress: 0.5, params: {} } } }])
    const lease = log.indexOf('lease:webgl2:mask')
    expect(lease).toBeGreaterThanOrEqual(0)
    expect(log.indexOf('gl:createTexture')).toBeGreaterThan(lease)
    comp.setSlideIndex(3)
    ws.pool.limit = ws.used()
    comp.frame([{ handle: a, state: { ...s, mask: { kind: 'dissolve', progress: 0.25, params: {} } } }])
    expect(sink.list().some(d => d.code === 'compositor-workspace' && d.slideIndex === 3)).toBe(true)
  })

  it('dispose returns every byte it leased', () => {
    const fake = fakeGL()
    current = fake
    const ws = makeWorkspace(1 << 30)
    const comp = createWebGL2Compositor(document.createElement('canvas'), ws)!
    layerOn(comp, ws)
    const s = { visible: true, opacity: 1, transform: [1, 0, 0, 1, 0, 0] as const }
    comp.frame([{ handle: { id: 1 }, state: { ...s, mask: { kind: 'wipe', progress: 0.5, params: {} } } }])
    expect(ws.used()).toBeGreaterThan(0)
    comp.dispose()
    expect(ws.used()).toBe(0)
  })
})

// ---------------------------------------------------------------------------
// The backend manager: context loss (round 2)
// ---------------------------------------------------------------------------
describe('createBackendManager', () => {
  it('starts on WebGL2 and mounts its canvas in the host', () => {
    const h = harness()
    expect(h.mgr.current.backend).toBe('webgl2')
    expect(h.mgr.state).toBe('webgl2')
    expect(h.glCanvas()).not.toBeNull()
    expect(h.liveCanvas()).toBeNull()
  })

  it('without WebGL2 starts on canvas2d with a diagnostic', () => {
    const h = harness({ gl: false })
    expect(h.mgr.current.backend).toBe('canvas2d')
    expect(h.mgr.state).toBe('canvas2d')
    expect(h.glCanvas()).toBeNull()
    expect(h.liveCanvas()).not.toBeNull()
    expect(h.sink.list().some(d => d.code === 'backend-fallback')).toBe(true)
  })

  it('loss mounts a 2-D canvas in the same task and preventDefault is called', () => {
    const h = harness()
    h.mgr.current.setViewport({ matrix: [1, 0, 0, 1, 0, 0], deviceWidth: 320, deviceHeight: 180 })
    const e = h.lose()
    // No await: the event handler did it all synchronously.
    expect(e.defaultPrevented).toBe(true)
    const live = h.liveCanvas()
    expect(live).toBe(h.output)
    expect(h.glCanvas()).not.toBeNull()
    // The lost WebGL canvas stacks over the output: it is hidden while live state paints beneath it.
    expect(h.glCanvas()!.style.visibility).toBe('hidden')
    expect(h.paintLive).toHaveBeenCalledTimes(1)
    expect(h.paintLive.mock.calls[0][0].canvas).toBe(live)
    expect(h.mgr.state).toBe('lost')
    // Every frame while lost paints live state on the 2-D canvas.
    h.mgr.current.frame([])
    h.mgr.current.frame([])
    expect(h.paintLive).toHaveBeenCalledTimes(3)
    expect(h.sink.list().some(d => d.code === 'context-lost')).toBe(true)
  })

  it('restore at 2999 ms recreates programs, buffers, VAOs, framebuffers and textures and switches back', async () => {
    const h = harness()
    h.mgr.current.setViewport({ matrix: [1, 0, 0, 1, 0, 0], deviceWidth: 64, deviceHeight: 64 })
    expect('id' in layerOn(h.mgr.current, h.ws)).toBe(true)
    const programs = h.fake.count('createProgram'), buffers = h.fake.count('createBuffer'), vaos = h.fake.count('createVertexArray')
    const textures = h.fake.count('createTexture'), framebuffers = h.fake.count('createFramebuffer')
    h.lose()
    h.advance(2999)
    expect(h.mgr.state).toBe('lost')
    h.log.length = 0
    h.restore()
    // Programs, buffers and vertex arrays come back at once, with GL state.
    expect(h.fake.count('createProgram')).toBe(programs + 1)
    expect(h.fake.count('createBuffer')).toBe(buffers + 1)
    expect(h.fake.count('createVertexArray')).toBe(vaos + 1)
    // The drawing buffer (the output framebuffer) is charged again; no offscreen framebuffer class exists yet.
    expect(h.log).toContain('lease:webgl2:framebuffer')
    expect(h.fake.count('createFramebuffer')).toBe(framebuffers)
    // Layer textures are re-rasterised and re-uploaded by the caller.
    expect(h.rebuilt).toEqual(['webgl2'])
    expect(h.mgr.state).toBe('lost')
    await h.settle()
    expect(h.fake.count('createTexture')).toBe(textures + 1)
    expect(h.mgr.state).toBe('webgl2')
    expect(h.mgr.current.backend).toBe('webgl2')
    expect(h.liveCanvas()).toBeNull()
    // The restored backend got the viewport back.
    h.fake.calls.length = 0
    const painted = h.paintLive.mock.calls.length
    h.mgr.current.frame([])
    expect(h.fake.calls.find(c => c.name === 'viewport')?.args).toEqual([0, 0, 64, 64])
    expect(h.paintLive).toHaveBeenCalledTimes(painted)
    // The deadline callback, when it fires later, changes nothing.
    h.advance(10_000)
    expect(h.mgr.state).toBe('webgl2')
  })

  it('no restore by 3000 ms switches to canvas2d; a later restore is ignored', async () => {
    const h = harness()
    h.lose()
    const gl = h.glCanvas()!
    h.advance(2999)
    expect(h.mgr.state).toBe('lost')
    h.advance(1)
    expect(h.mgr.state).toBe('canvas2d')
    expect(h.sink.list().some(d => d.code === 'backend-fallback' && d.effect === 'restore-deadline')).toBe(true)
    // The WebGL canvas is removed; the 2-D canvas carries on painting live state until the layers are built.
    expect(h.glCanvas()).toBeNull()
    expect(h.rebuilt).toEqual(['canvas2d'])
    await h.settle()
    expect(h.mgr.current.backend).toBe('canvas2d')
    const programs = h.fake.count('createProgram')
    h.restore(gl)
    await h.settle()
    expect(h.fake.count('createProgram')).toBe(programs)
    expect(h.rebuilt).toEqual(['canvas2d'])
    expect(h.mgr.state).toBe('canvas2d')
  })

  it('a restore that arrives at 3000 ms, before the timer fires, is also too late', async () => {
    const h = harness()
    h.lose()
    const gl = h.glCanvas()!
    const programs = h.fake.count('createProgram')
    // Move the clock without running the queue: the restore event beats the timer.
    h.jump(RESTORE_DEADLINE_MS)
    h.restore(gl)
    expect(h.fake.count('createProgram')).toBe(programs)
    expect(h.mgr.state).toBe('canvas2d')
  })

  it('pausing the show does not stretch the deadline', () => {
    const h = harness()
    const show = new ManualClock()
    h.lose()
    show.pause()
    // A throttled scheduler may fire early; the deadline is measured on now(), so nothing switches yet.
    h.advance(1000)
    h.fireEarly()
    expect(h.mgr.state).toBe('lost')
    show.advance(5000)
    expect(show.now()).toBe(0)
    h.advance(2000)
    expect(h.mgr.state).toBe('canvas2d')
  })

  it('a third loss pins canvas2d', async () => {
    const h = harness()
    for (let i = 0; i < 2; i++) {
      h.lose()
      h.advance(100)
      h.restore()
      await h.settle()
      expect(h.mgr.state).toBe('webgl2')
    }
    const e = h.lose()
    expect(e.defaultPrevented).toBe(true)
    expect(h.mgr.state).toBe('canvas2d')
    expect(h.sink.list().some(d => d.code === 'backend-fallback' && d.effect === 'repeated-loss')).toBe(true)
    await h.settle()
    expect(h.mgr.current.backend).toBe('canvas2d')
    const programs = h.fake.count('createProgram')
    h.restore(h.glCanvas() ?? undefined)
    await h.settle()
    expect(h.fake.count('createProgram')).toBe(programs)
    expect(h.mgr.state).toBe('canvas2d')
  })

  it('exit during restore wait ignores the restore', async () => {
    const h = harness()
    h.lose()
    const gl = h.glCanvas()!
    h.mgr.dispose()
    expect(h.host.querySelector('canvas')).toBeNull()
    const programs = h.fake.count('createProgram')
    h.restore(gl)
    h.advance(5000)
    await h.settle()
    expect(h.fake.count('createProgram')).toBe(programs)
    expect(h.deps.rebuildLayers).not.toHaveBeenCalled()
    expect(h.host.querySelector('canvas')).toBeNull()
    // The manager never releases the permanent reserve: the owner does.
    expect(h.ws.pool.used).toBe(h.terminal.outputBytes + (1 << 20))
  })

  it('a loss while the restored layers are still building starts a fresh wait', async () => {
    const h = harness()
    let open!: () => void
    h.rebuildGate.hold = new Promise<void>(r => { open = r })
    h.lose()
    h.restore()
    h.lose()
    open()
    await h.settle()
    // The first rebuild's resolution is stale: still lost, still painting live.
    expect(h.mgr.state).toBe('lost')
    expect(h.liveCanvas()).not.toBeNull()
  })

  it('restoration re-leases every recreated resource; a refusal drops to canvas2d with a diagnostic', async () => {
    const h = harness()
    h.mgr.current.setViewport({ matrix: [1, 0, 0, 1, 0, 0], deviceWidth: 64, deviceHeight: 64 })
    layerOn(h.mgr.current, h.ws)
    h.mgr.current.frame([{ handle: { id: 1 }, state: { visible: true, opacity: 1, transform: [1, 0, 0, 1, 0, 0], mask: { kind: 'wipe', progress: 0.5, params: {} } } }])
    const reserve = h.terminal.outputBytes + (1 << 20)
    expect(h.ws.used()).toBeGreaterThan(reserve)
    h.lose()
    // A lost context holds nothing: every GPU lease is returned.
    expect(h.ws.used()).toBe(reserve)
    h.log.length = 0
    h.restore()
    await h.settle()
    expect(h.log).toEqual(expect.arrayContaining(['lease:webgl2:framebuffer', 'lease:webgl2:buffer', 'grow:layer']))
    expect(h.mgr.state).toBe('webgl2')

    // Second loss; this time the workspace cannot take the drawing buffer back.
    h.lose()
    h.ws.pool.limit = h.ws.used() + 16
    h.restore()
    expect(h.mgr.state).toBe('canvas2d')
    expect(h.sink.list().some(d => d.code === 'backend-fallback' && d.effect === 'restore-workspace')).toBe(true)
    expect(h.glCanvas()).toBeNull()
  })

  it('the live 2-D canvas fits inside the terminal reserve: mounting it needs no admission', () => {
    const h = harness()
    h.mgr.current.setViewport({ matrix: [4, 0, 0, 4, 0, 0], deviceWidth: 1280, deviceHeight: 720 })
    h.log.length = 0
    h.lose()
    const live = h.liveCanvas()!
    expect(live.width * live.height * 4).toBeLessThanOrEqual(h.terminal.outputBytes)
    expect(h.log.filter(l => l.startsWith('lease:'))).toEqual([])
  })

  it('PR #18 review P2: a Canvas 2D start draws on the reserved output canvas and creates no canvas of its own', () => {
    const h = harness()
    const created = vi.spyOn(document, 'createElement')
    const canvases = () => created.mock.calls.flatMap((c, i) => c[0] === 'canvas' ? [created.mock.results[i].value as HTMLCanvasElement] : [])
    for (const opts of [{ force: 'canvas2d' as const }, {}]) {
      if (!('force' in opts)) current = null // WebGL2 unavailable
      const host = document.createElement('div')
      h.output.after(host)
      const before = canvases().length
      const mgr = createBackendManager(host, { ...h.deps, ...opts })
      expect(mgr.current.backend).toBe('canvas2d')
      // Forced: nothing is created. Unavailable: only the WebGL probe canvas, never sized and already gone.
      const made = canvases().slice(before)
      expect(made.length).toBe('force' in opts ? 0 : 1)
      for (const c of made) expect([c.isConnected, c.width, c.height]).toEqual([false, 300, 150])
      expect(host.querySelector('canvas')).toBeNull()
      expect(h.output.dataset.slideshowBackend).toBe('canvas2d')
      mgr.dispose()
      host.remove()
    }
  })

  it('PR #18 review P2: context loss and the deadline fallback draw on the reserved output canvas, never a new one', async () => {
    const h = harness()
    h.mgr.current.setViewport({ matrix: [1, 0, 0, 1, 0, 0], deviceWidth: 320, deviceHeight: 180 })
    let open!: () => void
    h.rebuildGate.hold = new Promise<void>(r => { open = r })
    const created = vi.spyOn(document, 'createElement')
    const canvases = () => created.mock.calls.filter(c => c[0] === 'canvas').length
    h.lose()
    expect(h.mgr.state).toBe('lost')
    expect(h.liveCanvas()).toBe(h.output)
    expect(h.paintLive.mock.calls[0][0].canvas).toBe(h.output)
    h.advance(RESTORE_DEADLINE_MS)
    expect(h.mgr.state).toBe('canvas2d')
    expect(h.glCanvas()).toBeNull()
    expect(h.liveCanvas()).toBe(h.output)
    expect(canvases()).toBe(0)
    expect(h.host.querySelector('canvas')).toBeNull()
    open()
    await h.settle()
    expect(h.mgr.current.backend).toBe('canvas2d')
    // The output is the reserve's surface: its size is the stage's, never the fallback's.
    expect([h.output.width, h.output.height]).toEqual([320, 180])
    h.mgr.dispose()
    expect(h.output.dataset.slideshowBackend).toBeUndefined()
  })

  it('a forced canvas2d manager never creates a WebGL context', () => {
    const h = harness()
    const host = document.createElement('div')
    const mgr = createBackendManager(host, { ...h.deps, force: 'canvas2d' })
    expect(mgr.current.backend).toBe('canvas2d')
    expect(host.querySelector('canvas[data-slideshow-backend="webgl2"]')).toBeNull()
    mgr.dispose()
  })
})

// ---------------------------------------------------------------------------
// The live painter assembly (compile-checked)
// ---------------------------------------------------------------------------
describe('createLivePainter', () => {
  it('paints through paintDirect with terminal.scratch as the DirectAssets scratch', () => {
    const ws = makeWorkspace(1 << 30)
    const terminal = ws.reserveTerminal(64 * 64 * 4)!
    const slide: PptxSlide = { index: 0, widthEmu: 32 * 9525, heightEmu: 32 * 9525, shapes: [], background: '#123456' }
    const seen: DirectAssets[] = []
    const scene: LiveScene = { slide, state: new Map(), assets: { images: [], resolveFont: (f: string) => f } }
    const paint = createLivePainter(terminal, () => scene, (ctx, s, st, assets) => { seen.push(assets); paintDirect(ctx, s, st, assets) })
    const canvas = document.createElement('canvas'); canvas.width = 32; canvas.height = 32
    const ctx = canvas.getContext('2d')!
    paint(ctx)
    expect(seen).toHaveLength(1)
    // The compile-time half: the scratch slot is typed as the terminal's SubOwner.
    const scratch: SubOwner = seen[0].scratch
    expect(scratch).toBe(terminal.scratch)
    expect(seen[0].scale).toBe(1)
    expect([...ctx.getImageData(5, 5, 1, 1).data]).toEqual([0x12, 0x34, 0x56, 255])
    // No scene (between slides): nothing painted, nothing thrown.
    expect(() => createLivePainter(terminal, () => undefined)(ctx)).not.toThrow()
  })
})

// ---------------------------------------------------------------------------
// The parity comparator (the two-backend comparison itself runs in Chromium)
// ---------------------------------------------------------------------------
describe('frameDiff', () => {
  const frame = (w: number, h: number, fill: (x: number, y: number) => number[]) => {
    const d = new Uint8ClampedArray(w * h * 4)
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) d.set(fill(x, y), (y * w + x) * 4)
    return d
  }
  it('measures mean and worst-row absolute difference as fractions of 255', () => {
    const a = frame(4, 2, () => [0, 0, 0, 255])
    expect(frameDiff(a, a, 4, 2)).toEqual({ meanAbs: 0, rowMax: 0 })
    // One row fully off by 51 in red: that row's mean is 51/4/255 = 0.05; overall half that.
    const b = frame(4, 2, (_x, y) => (y === 1 ? [51, 0, 0, 255] : [0, 0, 0, 255]))
    const d = frameDiff(a, b, 4, 2)
    expect(d.rowMax).toBeCloseTo(0.05, 6)
    expect(d.meanAbs).toBeCloseTo(0.025, 6)
    expect(withinTolerance(d, { meanAbs: 0.02, rowMax: 0.10 })).toBe(false)
    expect(withinTolerance(d, { meanAbs: 0.08, rowMax: 0.30 })).toBe(true)
  })
  it('a vertically flipped frame fails the exact tolerance', () => {
    const a = frame(8, 8, (_x, y) => (y < 4 ? [255, 0, 0, 255] : [0, 0, 255, 255]))
    const flipped = frame(8, 8, (_x, y) => (y >= 4 ? [255, 0, 0, 255] : [0, 0, 255, 255]))
    expect(withinTolerance(frameDiff(a, flipped, 8, 8), { meanAbs: 0.02, rowMax: 0.10 })).toBe(false)
  })
})

/** The type of a manager's dependencies is the brief's, plus the optional `force`. */
export type _Deps = Parameters<typeof createBackendManager>[1]
