/**
 * Show accessibility (slideshow spec 9): the audience surface's polite live
 * region ("Slide n of m: {title}", announced once per change), an offscreen
 * list of the slide's **currently visible** text (title, then body; text
 * whose entrance has not run is absent, exited text is removed), and one
 * offscreen focusable button per trigger shape and hyperlink in authored
 * z-order ("Trigger: {name}", "Link: {text}"). Tab moves through them and
 * Enter or Space activates them, the same as a click. Focus is never trapped.
 *
 * Plain DOM through the surface's own document (realm-safe, like
 * surface.tsx). With copying disallowed the text stays readable (each item
 * is named by `aria-label`) but carries no selectable text, the same
 * pattern as the paged viewer's offscreen chart layer.
 */
import type { PptxShape, PptxSlide } from '../../pptx/types'
import { shapeKey } from '../render/plan'
import { underlayOf } from '../render/scene'
import { targetKey } from '../target'
import type { BuildOwnership, LinkSpec, RenderState, TargetKey } from '../types'

/** `run`: the linked text run a link target is on (PR #18 review 3: each linked run is its own target). */
export interface ShowTarget { spid: string; kind: 'trigger' | 'link'; label: string; run?: { p: number; r: number } }

export interface ShowA11y {
  /** Sets the live region's text when it differs from the last announcement. */
  announce(text: string): void
  /** The slide's visible text, in reading order. */
  setText(items: readonly string[]): void
  /** The trigger and link buttons, in z-order. */
  setTargets(targets: readonly ShowTarget[]): void
  readonly live: HTMLElement
  dispose(): void
}

const HIDDEN = 'position:absolute;width:1px;height:1px;margin:-1px;padding:0;overflow:hidden;clip-path:inset(50%);white-space:nowrap;border:0'

/** A target's stable id (its button's `data-target`): kind and spid, plus the run for a run link. */
const idOf = (t: ShowTarget) => `${t.kind}:${t.spid}${t.run ? `:${t.run.p}:${t.run.r}` : ''}`

export function createShowA11y(root: HTMLElement, options: { selectable: boolean; activate(target: ShowTarget): void }): ShowA11y {
  const doc = root.ownerDocument
  const live = doc.createElement('div')
  live.setAttribute('data-officeview-show-live', '')
  live.setAttribute('aria-live', 'polite')
  const list = doc.createElement('ul')
  list.setAttribute('data-officeview-show-text', '')
  list.setAttribute('aria-label', 'Slide text')
  list.style.cssText = HIDDEN + (options.selectable ? '' : ';user-select:none;-webkit-user-select:none')
  const targets = doc.createElement('div')
  targets.setAttribute('data-officeview-show-targets', '')
  targets.style.cssText = HIDDEN
  root.append(live, list, targets)
  let announced = ''
  let text: readonly string[] = []
  let shown: readonly ShowTarget[] = []
  const same = <T,>(a: readonly T[], b: readonly T[], eq: (x: T, y: T) => boolean) => a.length === b.length && a.every((x, i) => eq(x, b[i]))
  return {
    live,
    announce(t) {
      if (t === announced) return
      announced = t
      live.textContent = t
    },
    setText(items) {
      if (same(items, text, (a, b) => a === b)) return
      text = [...items]
      list.replaceChildren(...items.map((t) => {
        const li = doc.createElement('li')
        if (options.selectable) li.textContent = t
        else li.setAttribute('aria-label', t)
        return li
      }))
    },
    setTargets(next) {
      if (same(next, shown, (a, b) => idOf(a) === idOf(b) && a.label === b.label)) return
      shown = [...next]
      // Keep focus on the same target across a rebuild (a trigger that re-renders its list).
      const focused = doc.activeElement && targets.contains(doc.activeElement) ? (doc.activeElement as HTMLElement).dataset.target : undefined
      targets.replaceChildren(...next.map((t) => {
        const b = doc.createElement('button')
        b.type = 'button'
        b.textContent = `${t.kind === 'trigger' ? 'Trigger' : 'Link'}: ${t.label}`
        b.dataset.target = idOf(t)
        b.addEventListener('click', (e) => { e.stopPropagation(); options.activate(t) })
        b.addEventListener('keydown', (e) => {
          if (e.key !== 'Enter' && e.key !== ' ') return
          e.preventDefault()
          e.stopPropagation()
          options.activate(t)
        })
        return b
      }))
      if (focused) {
        // The same target keeps focus; if it is gone (the slide changed), focus returns to the show host so its keys (Esc) still work.
        const again = ([...targets.children] as HTMLElement[]).find((b) => b.dataset.target === focused)
        ;(again ?? (root.closest('[data-officeview-show]') as HTMLElement | null) ?? root).focus()
      }
    },
    dispose() {
      live.remove()
      list.remove()
      targets.remove()
    },
  }
}

// ---------------------------------------------------------------------------
// What the lists hold
// ---------------------------------------------------------------------------

export function shapeText(shape: PptxShape, keep?: (paragraph: number) => boolean): string {
  const paras = shape.textBody?.paragraphs ?? []
  return paras
    .map((p, i) => (keep && !keep(i) ? '' : p.runs.map((r) => r.text ?? '').join('')))
    .filter((t) => t.trim() !== '')
    .join('\n')
}

const isTitle = (shape: PptxShape) => shape.placeholder?.type === 'title' || shape.placeholder?.type === 'ctrTitle'

function shown(state: RenderState, key: TargetKey | undefined): boolean {
  if (key === undefined) return true
  const s = state.get(key)
  return !s || (s.visible && s.opacity > 0)
}

/** Every leaf with its ancestors' visibility (a hidden group hides its children), in paint order. */
function leaves(nodes: readonly PptxShape[], state: RenderState, out: PptxShape[] = []): PptxShape[] {
  for (const n of nodes) {
    if (!shown(state, shapeKey(n))) continue
    if (n.group) leaves(n.children ?? [], state, out)
    else out.push(n)
  }
  return out
}

/** The slide title: the first title placeholder's text, whatever its animation state. */
export function slideTitle(slide: PptxSlide): string {
  const title = leaves(slide.shapes, new Map()).find(isTitle)
  return title ? shapeText(title).replace(/\s+/g, ' ').trim() : ''
}

/**
 * The slide's currently visible text in reading order: title, then the rest
 * in paint order. A paragraph build exposes only its visible ranges (and the
 * paragraphs outside every range when the shape itself is visible).
 */
export function visibleText(slide: PptxSlide, state: RenderState, builds: readonly BuildOwnership[]): string[] {
  const byShape = new Map(builds.filter((b) => b.mode === 'paragraph').map((b) => [b.spid, b]))
  const out: Array<{ title: boolean; text: string }> = []
  for (const shape of leaves([...underlayOf(slide), ...slide.shapes], state)) {
    const spid = shape.source?.id
    const build = spid !== undefined ? byShape.get(spid) : undefined
    let text: string
    if (build) {
      const ranges = build.ranges.filter((r) => shown(state, r.key))
      const inAny = (i: number) => build.ranges.some((r) => i >= r.st && i <= r.end)
      text = shapeText(shape, (i) => !inAny(i) || ranges.some((r) => i >= r.st && i <= r.end))
    } else text = shapeText(shape)
    if (text) out.push({ title: isTitle(shape), text })
  }
  return [...out.filter((t) => t.title), ...out.filter((t) => !t.title)].map((t) => t.text)
}

/** Trigger and link targets in authored z-order (paint order), visible ones only; a trigger wins over a link on the same shape. */
export function showTargets(slide: PptxSlide, state: RenderState, triggers: readonly string[], links: readonly LinkSpec[]): ShowTarget[] {
  const trig = new Set(triggers)
  const linked = new Map<string, LinkSpec[]>()
  for (const l of links) {
    const list = linked.get(l.spid) ?? []
    // One target per shape link and per linked run (a run with two hlinkClicks is one target).
    if (!list.some((k) => k.run?.p === l.run?.p && k.run?.r === l.run?.r)) list.push(l)
    linked.set(l.spid, list)
  }
  const out: ShowTarget[] = []
  const walk = (nodes: readonly PptxShape[]) => {
    for (const n of nodes) {
      const spid = n.source?.id
      if (!shown(state, shapeKey(n))) continue
      if (spid !== undefined && (trig.has(spid) || linked.has(spid))) {
        const name = n.source?.name?.trim()
        const text = shapeText(n).replace(/\s+/g, ' ').trim()
        if (trig.has(spid)) out.push({ spid, kind: 'trigger', label: name || text || `shape ${spid}` })
        else for (const l of linked.get(spid)!) out.push({ spid, kind: 'link', label: linkText(n, l) || text || name || `shape ${spid}`, ...(l.run ? { run: l.run } : {}) })
      }
      if (n.group) walk(n.children ?? [])
    }
  }
  walk(slide.shapes)
  return out
}

function linkText(shape: PptxShape, link: LinkSpec): string {
  if (!link.run) return ''
  const p = shape.textBody?.paragraphs[link.run.p]
  return p?.runs[link.run.r]?.text?.trim() ?? ''
}

/** The paragraph-range key of a build range, for callers holding only spid and indices. */
export const rangeKey = (spid: string, st: number, end: number): TargetKey => targetKey({ spid, text: { kind: 'p', st, end } })
