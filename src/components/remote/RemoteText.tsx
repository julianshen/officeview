/**
 * Stage 3a text services in the remote view (spec §4.1): the Ctrl/Cmd+F
 * search bar, search-hit and selection boxes, pointer/keyboard selection and
 * copy. Every answer comes from the server (`RemoteOfficeDoc.search/select/
 * copy`); this module never indexes text. The overlay is empty,
 * percent-positioned boxes only; the one place document text enters the DOM
 * is the visually-hidden screen-reader layer (Task 8), and only when the
 * server grants `can.textA11y` (`RemoteOfficeDoc.a11y`).
 */
import { useEffect, useRef, useState, type CSSProperties, type KeyboardEvent, type PointerEvent as ReactPointerEvent, type ReactElement, type ReactNode, type RefObject } from 'react'
import { copyText } from '../../core/clipboard'
import type { HighlightRect } from '../../core/overlay'
import type { Manifest } from '../../stream/manifest'
import type { RemoteOfficeDoc, RemoteRange, RemoteSearchHit } from '../../stream/remote/remote-doc'
import { TEXT_LIMITS } from '../../stream/remote/text-wire'
import { SearchBar } from '../viewerControls'

/** A drag issues at most one `select` per this many ms (≤ 5 per second); the final position is always sent. */
export const SELECT_INTERVAL_MS = 200
/** Typing settles this long before a search is sent. */
export const SEARCH_DEBOUNCE_MS = 200
/** R6: the server answers queries of at most 256 UTF-16 units. */
const MAX_QUERY_UNITS = 256
/** Pointer travel (CSS px) below which a press is a click, not a drag. */
const DRAG_SLOP_PX = 3

/** The status-region message for a refused or failed copy. */
export function copyFailure(e: unknown, caps: Manifest['caps']['copy']): string {
  const { code, limit } = (e ?? {}) as { code?: unknown; limit?: unknown }
  const n = typeof limit === 'number' ? limit : undefined
  if (code === 'copy-too-large') {
    // A paragraph cap states its limit; the 4 MiB byte cap (or no paragraph cap) does not.
    return caps?.maxParagraphs != null && n !== undefined && n !== TEXT_LIMITS.copyBytes ? `Selection too large to copy (limit: ${n} paragraphs).` : 'Selection too large to copy.'
  }
  if (code === 'copy-quota-exhausted' && n !== undefined) return `Copy limit for this session reached (${n} paragraphs).`
  if (code === 'rate-limited') return 'Too many requests. Try again shortly.'
  return 'Copy failed.'
}

const A11Y_UNAVAILABLE = 'Text unavailable.'
/** A unit whose a11y fetch failed while it stays in view is retried once after this long. */
export const A11Y_RETRY_MS = 5000
/** The status-region message for a refused or failed a11y request. */
export function a11yFailure(e: unknown): string {
  const { code, limit } = (e ?? {}) as { code?: unknown; limit?: unknown }
  return code === 'copy-quota-exhausted' && typeof limit === 'number' ? `Copy limit for this session reached (${limit} paragraphs).` : A11Y_UNAVAILABLE
}

const searchFailure = (e: unknown): string => {
  const code = (e as { code?: unknown } | undefined)?.code
  return code === 'rate-limited' ? 'too many requests' : code === 'text-unavailable' ? 'search unavailable' : 'search failed'
}

const isAbort = (e: unknown) => (e as { name?: unknown } | undefined)?.name === 'AbortError'
const inField = (t: EventTarget) => t instanceof Element && !!t.closest('input, textarea, select, [contenteditable]')

export interface UnitBoxes { hits: HighlightRect[]; active: HighlightRect[]; sel: HighlightRect[] }
const NO_BOXES: UnitBoxes = { hits: [], active: [], sel: [] }

export interface RemoteText {
  /** The search bar (null while closed), to mount under a positioned, non-scrolling ancestor. */
  bar: ReactElement | null
  /** The copy status message, if any (replaces the notice in the status region). */
  message: string | undefined
  /** Handles the view's text shortcuts; true when the event was consumed. */
  onKeyDown(e: KeyboardEvent): boolean
  /** Starts a selection drag on unit `i`; undefined when copy is off (no handler is attached). */
  onPointerDown: ((i: number, e: ReactPointerEvent<HTMLElement>) => void) | undefined
  boxes(i: number): UnitBoxes
  /** Unit `i` entered (`true`) or left the viewport. Entering requests its a11y
   * text once per mode (again after a failure); a no-op request when a11y is off. */
  a11yVisible(i: number, visible: boolean): void
  /** Unit `i`'s screen-reader layer, or null when a11y is off. */
  layer(i: number): ReactElement | null
}

/** Visually hidden, still in the accessibility tree (never display:none or aria-hidden). */
const SR_ONLY: CSSProperties = { position: 'absolute', top: 0, left: 0, width: 1, height: 1, margin: -1, padding: 0, border: 0, overflow: 'hidden', clip: 'rect(0 0 0 0)', clipPath: 'inset(50%)', whiteSpace: 'nowrap' }

type A11yMode = Manifest['caps']['a11y']
/** One a11y period: a view lifetime in one mode. It ends on unmount, a document
 * change or a mode change (a resume, R3); nothing it started lands after that. */
interface A11yPeriod {
  mode: A11yMode
  signal: AbortSignal
  /** Units asked for (in flight or answered). */
  asked: Set<number>
  /** 'focused' `unit:k` reads in flight or done. */
  reading: Set<string>
  /** The one pending delayed retry per unit, and the units that already used it since entering view. */
  timers: Map<number, ReturnType<typeof setTimeout>>
  retried: Set<number>
}
/** What the layer shows, tagged with the mode it was fetched in. */
interface A11yState { mode: A11yMode; units: ReadonlyMap<number, ReadonlyArray<string | undefined>>; texts: ReadonlyMap<string, string> }

/**
 * Search, selection and copy state for one remote document. `root` is the
 * scroller: it takes focus on a drag and when the search bar closes, and the
 * unit wrappers (`[data-unit-wrap]`) are found inside it.
 */
export function useRemoteText(doc: RemoteOfficeDoc, root: RefObject<HTMLElement | null>): RemoteText {
  // A resume may replace the manifest (R3: copy and a11y may change): re-render and re-derive.
  const [, setManifestRev] = useState(0)
  useEffect(() => doc.onManifestChange?.(() => setManifestRev((n) => n + 1)), [doc])
  const can = doc.manifest.can
  const units = doc.manifest.units.list

  // ── search ──
  const [open, setOpen] = useState(false)
  const [query, setQuery] = useState('')
  const [hits, setHits] = useState<RemoteSearchHit[]>([])
  const [done, setDone] = useState<{ total: number; truncated: boolean }>()
  const [failure, setFailure] = useState<string>()
  const [active, setActive] = useState(0)
  const [nav, setNav] = useState(0)
  const inputRef = useRef<HTMLInputElement | null>(null)
  const searchAc = useRef<AbortController | undefined>(undefined)
  useEffect(() => {
    // A query change aborts the search in flight; an empty query clears without a request.
    searchAc.current?.abort(); searchAc.current = undefined
    setFailure(undefined)
    if (!open || query === '') { setHits([]); setDone(undefined); return }
    if (query.length > MAX_QUERY_UNITS) { setHits([]); setDone(undefined); setFailure('query too long'); return }
    const t = setTimeout(() => {
      const ac = new AbortController()
      searchAc.current = ac
      setHits([]); setDone(undefined); setActive(0)
      // Each onHits call carries every hit so far: replace, in document order
      // (unit, then arrival within the unit — Array#sort is stable).
      doc.search(query, { signal: ac.signal, onHits: (h) => { if (!ac.signal.aborted) setHits([...h].sort((a, b) => a.unit - b.unit)) } })
        .then((d) => { if (!ac.signal.aborted) setDone(d) }, (e) => { if (!ac.signal.aborted && !isAbort(e)) setFailure(searchFailure(e)) })
    }, SEARCH_DEBOUNCE_MS)
    return () => clearTimeout(t)
  }, [doc, open, query])
  useEffect(() => { if (open) inputRef.current?.focus() }, [open])
  const count = done ? done.total : hits.length
  const current = hits.length ? Math.min(active, hits.length - 1) : 0
  const step = (d: 1 | -1) => {
    if (!hits.length) return
    setActive((current + d + hits.length) % hits.length)
    setNav((n) => n + 1)
  }
  useEffect(() => {
    // next/prev bring the active hit (and so its unit) into view.
    if (!nav) return
    const el = root.current?.querySelector('[data-hit-active]') as HTMLElement | null
    el?.scrollIntoView?.({ block: 'center', inline: 'nearest' })
  }, [nav, root])
  const close = () => {
    setOpen(false); setQuery('')
    root.current?.focus({ preventScroll: true })
  }

  // ── selection ──
  const [sel, setSel] = useState<RemoteSearchHit[]>([])
  const [message, setMessage] = useState<string>()
  const range = useRef<RemoteRange | undefined>(undefined)
  /** Bumped by every select sent, a cleared selection and the end of a view
   * lifetime: only the newest select answer draws (a drag keeps live feedback). */
  const selGen = useRef(0)
  /** Bumped by every selection change and the end of a view lifetime: a copy
   * started for an older selection neither writes the clipboard nor sets a message. */
  const selVersion = useRef(0)
  /** The current view lifetime's controller (created in the effect, so a StrictMode
   * replay or a document change starts a fresh one); undefined between lifetimes. */
  const life = useRef<AbortController | undefined>(undefined)
  const selAc = useRef<AbortController | undefined>(undefined)
  /** The select limiter, shared by drags and Ctrl/Cmd+A: one request per SELECT_INTERVAL_MS,
   * the latest range riding as the trailing request. */
  const limiter = useRef<{ last: number; timer?: ReturnType<typeof setTimeout> }>({ last: -Infinity })
  const copying = useRef(false)
  // ── a11y text (Task 8) ──
  const a11yMode: A11yMode = can.textA11y ? doc.manifest.caps.a11y : 'off'
  /** Per unit: its paragraphs ('unit': their text; 'focused': placeholders to read) and
   * the 'focused' texts read so far by `unit:k` (cached for the period). */
  const [a11yState, setA11yState] = useState<A11yState>(() => ({ mode: a11yMode, units: new Map(), texts: new Map() }))
  const a11y = useRef<A11yPeriod | undefined>(undefined)
  /** The units in view now, for the current document. */
  const shown = useRef<{ doc: RemoteOfficeDoc; units: Set<number> }>({ doc, units: new Set() })
  const shownUnits = () => { if (shown.current.doc !== doc) shown.current = { doc, units: new Set() }; return shown.current.units }
  /** The `unit:k` whose <p> takes focus when it mounts (its button had focus when activated). */
  const focusAfter = useRef<string | undefined>(undefined)
  const drag = useRef<{ anchor: { unit: number; x: number; y: number }; unit: number; sx: number; sy: number; moved: boolean; detach(): void } | undefined>(undefined)
  const sendSelect = () => {
    const r = range.current, ac = selAc.current
    if (!r || !ac || ac.signal.aborted) return
    const gen = ++selGen.current
    // A newer select supersedes an older one ('closed'); only the current generation draws.
    doc.select(r, ac.signal).then((s) => { if (gen === selGen.current && !ac.signal.aborted) setSel(s.units) }, () => {})
  }
  const requestSelect = () => {
    const l = limiter.current
    if (l.timer) return // the pending trailing request sends the latest range
    const wait = l.last + SELECT_INTERVAL_MS - Date.now()
    if (wait <= 0) { l.last = Date.now(); sendSelect(); return }
    l.timer = setTimeout(() => { l.timer = undefined; l.last = Date.now(); sendSelect() }, wait)
  }
  const cancelPending = () => { const l = limiter.current; if (l.timer) { clearTimeout(l.timer); l.timer = undefined } }
  /** A selection change: earlier select/copy answers no longer apply. */
  const select = (r: RemoteRange) => { range.current = r; selVersion.current++; setMessage(undefined); requestSelect() }
  const clearSelection = () => {
    range.current = undefined; selGen.current++; selVersion.current++; cancelPending()
    selAc.current?.abort(); selAc.current = life.current && new AbortController()
    setSel([])
  }
  useEffect(() => {
    const ac = new AbortController()
    life.current = ac; selAc.current = new AbortController()
    return () => {
      // End of this lifetime (unmount, StrictMode replay, document change):
      // nothing it started may draw, write the clipboard or set a message.
      ac.abort(); selAc.current?.abort()
      life.current = undefined; selAc.current = undefined
      selGen.current++; selVersion.current++; cancelPending(); copying.current = false
      drag.current?.detach(); drag.current = undefined
      range.current = undefined
      searchAc.current?.abort(); searchAc.current = undefined
      setSel([]); setMessage(undefined); setHits([]); setDone(undefined)
    }
  }, [doc])
  // Copy turned off by a resume: the selection goes and nothing in flight may copy.
  const canCopy = can.copy
  useEffect(() => {
    if (canCopy) return
    drag.current?.detach(); drag.current = undefined
    clearSelection()
  }, [canCopy])
  /** Client px → page px of unit `i` at scale 1 (`x * unit.w / rect.width`), clamped to the unit. */
  const pointIn = (i: number, cx: number, cy: number) => {
    const wrap = root.current?.querySelector(`[data-unit-wrap="${i}"]`)
    const r = wrap?.getBoundingClientRect(), u = units[i]!
    const sx = r && r.width ? u.w / r.width : 1, sy = r && r.height ? u.h / r.height : 1
    const clamp = (v: number, max: number) => Math.min(max, Math.max(0, v))
    return { unit: i, x: clamp((cx - (r?.left ?? 0)) * sx, u.w), y: clamp((cy - (r?.top ?? 0)) * sy, u.h) }
  }
  /** The unit under the pointer, else the last one the drag was over. */
  const unitAt = (cx: number, cy: number, fallback: number) => {
    const hit = typeof document.elementFromPoint === 'function' ? document.elementFromPoint(cx, cy)?.closest('[data-unit-wrap]') : null
    return hit && root.current?.contains(hit) ? Number(hit.getAttribute('data-unit-wrap')) : fallback
  }
  const onPointerDown = can.copy ? (i: number, e: ReactPointerEvent<HTMLElement>) => {
    if (e.button !== 0) return
    // No native text selection or image drag; the view takes focus for Ctrl/Cmd+A/C.
    e.preventDefault()
    root.current?.focus({ preventScroll: true })
    drag.current?.detach()
    const to = (ev: PointerEvent) => {
      const d = drag.current!
      d.unit = unitAt(ev.clientX, ev.clientY, d.unit)
      select({ anchor: d.anchor, focus: pointIn(d.unit, ev.clientX, ev.clientY) })
    }
    const move = (ev: PointerEvent) => {
      const d = drag.current!
      if (!d.moved && Math.hypot(ev.clientX - d.sx, ev.clientY - d.sy) < DRAG_SLOP_PX) return
      d.moved = true
      to(ev)
    }
    const up = (ev: PointerEvent) => {
      const d = drag.current!
      // A click without movement clears the selection; a drag's final position
      // always goes (as the trailing request when the limiter is closed).
      if (!d.moved) clearSelection()
      else if (ev.type === 'pointerup') to(ev)
      d.detach()
      drag.current = undefined
    }
    const detach = () => {
      window.removeEventListener('pointermove', move)
      window.removeEventListener('pointerup', up)
      window.removeEventListener('pointercancel', up)
    }
    drag.current = { anchor: pointIn(i, e.clientX, e.clientY), unit: i, sx: e.clientX, sy: e.clientY, moved: false, detach }
    window.addEventListener('pointermove', move)
    window.addEventListener('pointerup', up)
    window.addEventListener('pointercancel', up)
  } : undefined

  const copy = () => {
    const r = range.current, lifetime = life.current
    // One copy at a time: a Ctrl/Cmd+C while one is in flight sends nothing.
    if (!r || !lifetime || copying.current) return
    copying.current = true
    const version = selVersion.current
    const live = () => !lifetime.signal.aborted && version === selVersion.current
    const settle = () => { if (!lifetime.signal.aborted) copying.current = false }
    doc.copy(r).then(
      async (text) => {
        if (!live()) return
        // The helper re-checks before the native write and before its fallback.
        const out = await copyText(text, { allowed: live, signal: lifetime.signal })
        if (live()) setMessage(out === 'failed' ? 'Copy failed.' : undefined)
      },
      (e) => { if (live()) setMessage(copyFailure(e, doc.manifest.caps.copy)) },
    ).finally(settle)
  }

  const a11yRequest = (p: A11yPeriod, i: number) => {
    if (p.mode === 'off' || p.signal.aborted || p.asked.has(i)) return
    p.asked.add(i)
    Promise.resolve().then(() => doc.a11y(i)).then(
      (ps) => {
        if (p.signal.aborted) return
        setA11yState((s) => (s.mode !== p.mode ? s : { ...s, units: new Map(s.units).set(i, ps.map((x) => (p.mode === 'unit' ? x.text ?? '' : undefined))) }))
        setMessage((m) => (m === A11Y_UNAVAILABLE ? undefined : m))
      },
      (e) => {
        if (p.signal.aborted) return
        // The layer stays empty; the next visibility retries, and a unit that stays in view once more after A11Y_RETRY_MS.
        p.asked.delete(i)
        setMessage(a11yFailure(e))
        if (!shownUnits().has(i) || p.retried.has(i) || p.timers.has(i)) return
        p.retried.add(i)
        p.timers.set(i, setTimeout(() => { p.timers.delete(i); if (shownUnits().has(i)) a11yRequest(p, i) }, A11Y_RETRY_MS))
      },
    )
  }
  useEffect(() => {
    // A new period per document and mode: the previous layer contents and cached texts are dropped,
    // its answers and timers die with it, and the units in view are fetched again in the new mode.
    const ac = new AbortController()
    const p: A11yPeriod = { mode: a11yMode, signal: ac.signal, asked: new Set(), reading: new Set(), timers: new Map(), retried: new Set() }
    a11y.current = p
    focusAfter.current = undefined
    setA11yState({ mode: a11yMode, units: new Map(), texts: new Map() })
    for (const i of shownUnits()) a11yRequest(p, i)
    return () => {
      ac.abort()
      for (const t of p.timers.values()) clearTimeout(t)
      if (a11y.current === p) a11y.current = undefined
    }
  }, [doc, a11yMode])
  const a11yVisible = (i: number, visible: boolean) => {
    const units = shownUnits(), p = a11y.current
    if (!visible) {
      units.delete(i)
      if (p) { const t = p.timers.get(i); if (t) { clearTimeout(t); p.timers.delete(i) } }
      return
    }
    if (units.has(i)) return
    units.add(i)
    if (!p) return
    p.retried.delete(i)
    a11yRequest(p, i)
  }
  const a11yRead = (i: number, k: number, button: HTMLElement) => {
    const p = a11y.current, key = `${i}:${k}`
    // Once per paragraph: an activation while the read is in flight, or after it landed, sends nothing.
    if (!p || p.mode !== 'focused' || p.signal.aborted || p.reading.has(key)) return
    p.reading.add(key)
    const failed = (e: unknown) => { if (p.signal.aborted) return; p.reading.delete(key); setMessage(a11yFailure(e)) }
    Promise.resolve().then(() => doc.a11y(i, k)).then((ps) => {
      if (p.signal.aborted) return
      const text = ps[k]?.text
      if (text === undefined) return failed(undefined)
      if (typeof document !== 'undefined' && document.activeElement === button) focusAfter.current = key
      setA11yState((s) => (s.mode !== p.mode ? s : { ...s, texts: new Map(s.texts).set(key, text) }))
    }, failed)
  }
  const layer = (i: number): ReactElement | null => {
    if (a11yMode === 'off') return null
    // Contents fetched in another mode never render (a resume changed it; the new period refetches).
    const current = a11yState.mode === a11yMode
    const ps = current ? a11yState.units.get(i) ?? [] : []
    return (
      <div data-a11y-layer="" data-ov-remote-chrome="" role="group" aria-label={`Page ${i + 1}`} style={SR_ONLY}>
        {a11yMode === 'unit' ? ps.map((t, k) => <p key={k}>{t}</p>) : ps.map((_, k) => {
          const key = `${i}:${k}`, text = a11yState.texts.get(key)
          return text === undefined
            ? <button key={k} type="button" onClick={(e) => a11yRead(i, k, e.currentTarget)}>{`Paragraph ${k + 1} of ${ps.length}. Press Enter to read.`}</button>
            : <p key={k} tabIndex={0} ref={(el) => { if (el && focusAfter.current === key) { focusAfter.current = undefined; el.focus() } }}>{text}</p>
        })}
      </div>
    )
  }

  const onKeyDown = (e: KeyboardEvent): boolean => {
    if (e.key === 'Escape' && open) { e.preventDefault(); close(); return true }
    if (!(e.metaKey || e.ctrlKey) || e.altKey) return false
    const k = e.key.toLowerCase()
    if (k === 'f' && can.search) {
      e.preventDefault()
      if (open) { inputRef.current?.focus(); inputRef.current?.select() } else setOpen(true)
      return true
    }
    if (!can.copy || inField(e.target)) return false
    if (k === 'a') { e.preventDefault(); select({ all: true }); return true }
    if (k === 'c' && range.current) { e.preventDefault(); copy(); return true }
    return false
  }

  const bar = open ? (
    <SearchBar inputRef={inputRef} query={query} onQuery={setQuery} count={count} active={current} indexing={false}
      onPrev={() => step(-1)} onNext={() => step(1)}
      {...(done?.truncated ? { countLabel: `${done.total}+` } : {})} {...(failure ? { failure } : {})} />
  ) : null

  const byUnit = new Map<number, UnitBoxes>()
  const at = (u: number) => { let b = byUnit.get(u); if (!b) byUnit.set(u, b = { hits: [], active: [], sel: [] }); return b }
  hits.forEach((h, k) => { const b = at(h.unit); for (const r of h.rects) { b.hits.push(r); if (k === current) b.active.push(r) } })
  for (const s of sel) at(s.unit).sel.push(...s.rects)
  return { bar, message, onKeyDown, onPointerDown, boxes: (i) => byUnit.get(i) ?? NO_BOXES, a11yVisible, layer }
}

const pct = (v: number, of: number) => `${(v * 100) / of}%`
function Box({ r, w, h, kind, active }: { r: HighlightRect; w: number; h: number; kind: 'hit' | 'sel'; active?: boolean }): ReactElement {
  const style: CSSProperties = {
    position: 'absolute', left: pct(r.x, w), top: pct(r.y, h), width: pct(r.width, w), height: pct(r.height, h), pointerEvents: 'none',
    ...(kind === 'sel' ? { background: 'var(--ov-color-text-selection, rgba(64,128,255,.35))' }
      : { background: 'var(--ov-color-find, rgba(255,214,0,.45))', ...(active ? { outline: '1.5px solid var(--ov-color-find-active, #ca5010)' } : {}) }),
  }
  return kind === 'sel' ? <div data-sel="" aria-hidden="true" style={style} />
    : <div data-hit="" {...(active ? { 'data-hit-active': '' } : {})} aria-hidden="true" style={style} />
}

/**
 * One unit: a positioned wrapper sized exactly like its canvas (`children`),
 * with the empty overlay boxes on top and, beside them, the screen-reader
 * layer (`layer`, absolutely positioned: the wrapper's box stays the canvas's).
 */
export function UnitWrap({ index, w, h, boxes, onPointerDown, layer, children }: { index: number; w: number; h: number; boxes: UnitBoxes; onPointerDown?: (i: number, e: ReactPointerEvent<HTMLElement>) => void; layer?: ReactNode; children: ReactNode }): ReactElement {
  const has = boxes.hits.length + boxes.sel.length > 0
  return (
    <div data-unit-wrap={index} style={{ position: 'relative', width: w, maxWidth: '100%' }}
      {...(onPointerDown ? { onPointerDown: (e: ReactPointerEvent<HTMLElement>) => onPointerDown(index, e) } : {})}>
      {children}
      {has ? (
        <div data-unit-overlay="" data-ov-remote-chrome="" aria-hidden="true" style={{ position: 'absolute', inset: 0, pointerEvents: 'none', overflow: 'hidden' }}>
          {boxes.hits.map((r, k) => <Box key={`h${k}`} r={r} w={w} h={h} kind="hit" active={boxes.active.includes(r)} />)}
          {boxes.sel.map((r, k) => <Box key={`s${k}`} r={r} w={w} h={h} kind="sel" />)}
        </div>
      ) : null}
      {layer}
    </div>
  )
}
