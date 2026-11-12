/**
 * Stage 3a text services in the remote view (spec §4.1): the Ctrl/Cmd+F
 * search bar, search-hit and selection boxes, pointer/keyboard selection and
 * copy. Every answer comes from the server (`RemoteOfficeDoc.search/select/
 * copy`); this module never indexes text and never puts document text in the
 * DOM — the overlay is empty, percent-positioned boxes only.
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
}

/**
 * Search, selection and copy state for one remote document. `root` is the
 * scroller: it takes focus on a drag and when the search bar closes, and the
 * unit wrappers (`[data-unit-wrap]`) are found inside it.
 */
export function useRemoteText(doc: RemoteOfficeDoc, root: RefObject<HTMLElement | null>): RemoteText {
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
  useEffect(() => () => searchAc.current?.abort(), [])
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
  const selGen = useRef(0)
  const selAc = useRef(new AbortController())
  const drag = useRef<{ anchor: { unit: number; x: number; y: number }; unit: number; sx: number; sy: number; moved: boolean; last: number; timer?: ReturnType<typeof setTimeout>; detach(): void } | undefined>(undefined)
  const sendSelect = () => {
    const r = range.current
    if (!r) return
    const gen = ++selGen.current
    // A newer select supersedes an older one ('closed'); only the newest answer draws.
    doc.select(r, selAc.current.signal).then((s) => { if (gen === selGen.current) setSel(s.units) }, () => {})
  }
  const clearSelection = () => {
    range.current = undefined; selGen.current++
    selAc.current.abort(); selAc.current = new AbortController()
    setSel([])
  }
  useEffect(() => () => { selAc.current.abort(); drag.current?.detach() }, [])
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
    const move = (ev: PointerEvent) => {
      const d = drag.current!
      if (!d.moved && Math.hypot(ev.clientX - d.sx, ev.clientY - d.sy) < DRAG_SLOP_PX) return
      if (!d.moved) { d.moved = true; setMessage(undefined) }
      d.unit = unitAt(ev.clientX, ev.clientY, d.unit)
      range.current = { anchor: d.anchor, focus: pointIn(d.unit, ev.clientX, ev.clientY) }
      if (d.timer) return
      const wait = SELECT_INTERVAL_MS - (Date.now() - d.last)
      if (wait <= 0) { d.last = Date.now(); sendSelect(); return }
      d.timer = setTimeout(() => { d.timer = undefined; d.last = Date.now(); sendSelect() }, wait)
    }
    const up = (ev: PointerEvent) => {
      const d = drag.current!
      d.detach()
      drag.current = undefined
      // A click without movement clears the selection.
      if (!d.moved) { clearSelection(); return }
      if (ev.type === 'pointerup') { d.unit = unitAt(ev.clientX, ev.clientY, d.unit); range.current = { anchor: d.anchor, focus: pointIn(d.unit, ev.clientX, ev.clientY) } }
      sendSelect()
    }
    const detach = () => {
      const d = drag.current
      if (d?.timer) { clearTimeout(d.timer); d.timer = undefined }
      window.removeEventListener('pointermove', move)
      window.removeEventListener('pointerup', up)
      window.removeEventListener('pointercancel', up)
    }
    drag.current = { anchor: pointIn(i, e.clientX, e.clientY), unit: i, sx: e.clientX, sy: e.clientY, moved: false, last: 0, detach }
    window.addEventListener('pointermove', move)
    window.addEventListener('pointerup', up)
    window.addEventListener('pointercancel', up)
  } : undefined

  const copy = () => {
    const r = range.current
    if (!r) return
    doc.copy(r).then(
      async (text) => { setMessage((await copyText(text)) === 'failed' ? 'Copy failed.' : undefined) },
      (e) => setMessage(copyFailure(e, doc.manifest.caps.copy)),
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
    if (k === 'a') { e.preventDefault(); setMessage(undefined); range.current = { all: true }; sendSelect(); return true }
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
  return { bar, message, onKeyDown, onPointerDown, boxes: (i) => byUnit.get(i) ?? NO_BOXES }
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
 * with the empty overlay boxes on top. A per-unit layer (Task 8's
 * screen-reader text) is another child of this wrapper.
 */
export function UnitWrap({ index, w, h, boxes, onPointerDown, children }: { index: number; w: number; h: number; boxes: UnitBoxes; onPointerDown?: (i: number, e: ReactPointerEvent<HTMLElement>) => void; children: ReactNode }): ReactElement {
  const has = boxes.hits.length + boxes.sel.length > 0
  return (
    <div data-unit-wrap={index} style={{ position: 'relative', width: w, maxWidth: '100%' }}
      {...(onPointerDown ? { onPointerDown: (e: ReactPointerEvent<HTMLElement>) => onPointerDown(index, e) } : {})}>
      {children}
      {has ? (
        <div data-unit-overlay="" aria-hidden="true" style={{ position: 'absolute', inset: 0, pointerEvents: 'none', overflow: 'hidden' }}>
          {boxes.hits.map((r, k) => <Box key={`h${k}`} r={r} w={w} h={h} kind="hit" active={boxes.active.includes(r)} />)}
          {boxes.sel.map((r, k) => <Box key={`s${k}`} r={r} w={w} h={h} kind="sel" />)}
        </div>
      ) : null}
    </div>
  )
}
