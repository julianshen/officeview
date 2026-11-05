import { useEffect, useRef, useState, type CSSProperties, type KeyboardEvent as ReactKeyboardEvent, type ReactElement } from 'react'
import { logicalToViewport, type ViewportSize } from '../../../xlsx/view/scrollWindow'
import { clipBoxToPaneRegion, paneRegionClips, paneRegionKind } from '../../../xlsx/view/frozenPanes'
import type { ScrollWindow } from '../../../xlsx/view/contracts'
import { PX_PER_PT } from '../../../xlsx/pivots/index'
import { timelineLayout } from '../../../xlsx/pivots/slicers/paint'
import type { Box, TimelineLevel, TimelineModel } from '../../../xlsx/pivots/slicers/model'
import type { TimelineController } from '../../../xlsx/pivots/slicers/controller'

export interface TimelineLayerProps {
  model: TimelineModel
  box: Box
  window: ScrollWindow
  viewport: ViewportSize
  controller: TimelineController
  invoker?: HTMLElement
}

const hiddenButton: CSSProperties = { position: 'absolute', boxSizing: 'border-box', margin: 0, padding: 0, border: 0, background: 'transparent', color: 'transparent', overflow: 'hidden', cursor: 'pointer', pointerEvents: 'auto' }
const levels: readonly TimelineLevel[] = ['years', 'quarters', 'months', 'days']
const periodName = (period: { label: string; group?: string }): string => [period.label, period.group].filter(Boolean).join(' ')

/** Accessible, bounded hit targets over the canvas-painted timeline. Unmeasured drag and scroll stay disabled by S5-C. */
export function TimelineLayer({ model: savedModel, box, window, viewport, controller, invoker }: TimelineLayerProps): ReactElement | null {
  const [, tick] = useState(0)
  const root = useRef<HTMLDivElement>(null)
  const levelPicker = useRef<HTMLSelectElement>(null), clearButton = useRef<HTMLButtonElement>(null)
  const active = useRef(0), anchor = useRef<number | null>(null), draftEnd = useRef<number | null>(null)
  const pointerStart = useRef<number | null>(null), pointerDragged = useRef(false), ignoreClick = useRef(false)
  const [draftLabel, setDraftLabel] = useState('')
  useEffect(() => controller.subscribe(() => tick(value => value + 1)), [controller])
  const model = controller.model(savedModel.name) ?? savedModel
  const zoom = window.zoom
  const layout = timelineLayout(model, box, PX_PER_PT)
  const clips = paneRegionClips(window.frozen, window.logicalOrigin, zoom, viewport)
  const barAt = logicalToViewport(window, { x: layout.bar.x, y: layout.bar.y })
  const barBox = { x: barAt.x, y: barAt.y, width: layout.bar.width * zoom, height: layout.bar.height * zoom }
  // Worksheet-logical rectangles become viewport boxes clipped to the painted
  // bar (for periods: the layout keeps one lookahead period the canvas clips)
  // and to their pane region (parts scrolled under a frozen strip stay hidden).
  const place = (lx: number, ly: number, lw: number, lh: number, clipBar: boolean) => {
    const at = logicalToViewport(window, { x: lx, y: ly })
    let rect = { x: at.x, y: at.y, width: lw * zoom, height: lh * zoom }
    if (clipBar) {
      const x0 = Math.max(rect.x, barBox.x), y0 = Math.max(rect.y, barBox.y)
      const x1 = Math.min(rect.x + rect.width, barBox.x + barBox.width), y1 = Math.min(rect.y + rect.height, barBox.y + barBox.height)
      if (x1 <= x0 || y1 <= y0) return undefined
      rect = { x: x0, y: y0, width: x1 - x0, height: y1 - y0 }
    }
    const clipped = clipBoxToPaneRegion(clips, paneRegionKind(window.frozen, lx, ly), rect)
    if (!clipped) return undefined
    if (clipped.x >= viewport.width || clipped.y >= viewport.height || clipped.x + clipped.width <= 0 || clipped.y + clipped.height <= 0) return undefined
    return clipped
  }
  const periods = layout.periods.map(period => {
    const rect = place(period.x, layout.bar.y, period.width, layout.bar.height, true)
    if (!rect) return undefined
    return { start: period.start, index: period.index, selected: period.selected, name: periodName(period), rect }
  }).filter((entry): entry is { start: number; index: number; selected: boolean; name: string; rect: { x: number; y: number; width: number; height: number } } => entry !== undefined)
  if (!periods.length) return null
  const levelBox = layout.level ? place(layout.level.right - 76, layout.level.baseline - 20, 76, 20, false) : undefined
  const clearBox = place(box.x + Math.max(6, box.width - 138), box.y + 30, 62, 24, false)
  const trackBox = layout.track ? place(layout.track.x, layout.track.y, layout.track.width, layout.track.height, false) : undefined
  const visibleIndex = (index: number) => Math.max(0, Math.min(periods.length - 1, index))
  const announce = (message: string) => { root.current?.setAttribute('aria-label', `${model.caption ?? model.name} timeline. ${message}`); controller.report(message) }
  const commitRange = (a: number, b: number) => {
    const first = periods[visibleIndex(Math.min(a, b))]!, last = periods[visibleIndex(Math.max(a, b))]!
    void controller.selectTimeline(model.name, first.start, last.start, new AbortController().signal)
    anchor.current = null; draftEnd.current = null; setDraftLabel('')
  }
  const selectLevel = (level: TimelineLevel) => { void controller.setTimelineLevel(model.name, level, new AbortController().signal) }
  const pointIndex = (event: { target: EventTarget | null }) => {
    const button = event.target instanceof Element ? event.target.closest<HTMLButtonElement>('[data-timeline-index]') : null
    const index = button?.dataset.timelineIndex
    return index === undefined ? undefined : Number(index)
  }
  const focusGrid = () => (invoker ?? root.current?.closest('[data-officeview-sheet-pane]')?.querySelector<HTMLElement>('[role="grid"]'))?.focus()
  const lastPeriod = () => root.current?.querySelector<HTMLButtonElement>(`[data-timeline-index="${periods.length - 1}"]`)?.focus()
  const onPeriodTab = (event: ReactKeyboardEvent<HTMLButtonElement>) => {
    if (event.key !== 'Tab') return
    event.preventDefault()
    if (event.shiftKey) { focusGrid(); return }
    if (levelPicker.current) levelPicker.current.focus()
    else if (clearButton.current) clearButton.current.focus()
    else focusGrid()
  }
  const styleAt = (rect: { x: number; y: number; width: number; height: number }): CSSProperties => ({ ...hiddenButton, left: rect.x, top: rect.y, width: Math.max(0, rect.width), height: Math.max(0, rect.height) })
  const periodStart = (period: typeof periods[number]) => period.start
  return <div ref={root} role="group" aria-label={`${model.caption ?? model.name} timeline`} data-officeview-ui="timeline-layer" data-timeline-name={model.name}
    style={{ position: 'absolute', left: 0, top: 0, width: viewport.width, height: viewport.height, overflow: 'hidden', pointerEvents: 'none' }}
    onPointerDown={event => { pointerStart.current = pointIndex(event) ?? null; pointerDragged.current = false }}
    onPointerMove={event => { const index = pointIndex(event); if (pointerStart.current !== null && index !== undefined && index !== pointerStart.current) pointerDragged.current = true }}
    onPointerUp={() => {
      if (pointerDragged.current) { ignoreClick.current = true; announce('Timeline drag has not been measured and is unavailable.') }
      pointerStart.current = null; pointerDragged.current = false
    }}
    onPointerCancel={() => { pointerStart.current = null; pointerDragged.current = false; ignoreClick.current = true; announce('Timeline drag has not been measured and is unavailable.') }}>
    {periods.map((period, visibleAt) => {
      const committed = period.selected
      const drafted = anchor.current !== null && draftEnd.current !== null && visibleAt >= Math.min(anchor.current, draftEnd.current) && visibleAt <= Math.max(anchor.current, draftEnd.current)
      return <button key={`${period.start}:${period.index}`} type="button" data-timeline-index={visibleAt} tabIndex={-1} aria-label={`Select ${period.name}`}
        aria-pressed={committed} data-range-draft={drafted || undefined} style={styleAt(period.rect)}
        onFocus={() => { active.current = visibleAt }}
        onClick={() => {
          if (ignoreClick.current) { ignoreClick.current = false; return }
          active.current = visibleAt
          if (anchor.current !== null) commitRange(anchor.current, visibleAt)
          else void controller.selectTimeline(model.name, periodStart(period), periodStart(period), new AbortController().signal)
        }}
        onKeyDown={event => {
          if (event.key === 'Tab') onPeriodTab(event)
          else if (event.key === 'ArrowLeft' || event.key === 'ArrowRight') {
            event.preventDefault()
            const direction = event.key === 'ArrowRight' ? 1 : -1
            active.current = visibleIndex(visibleAt + direction)
            if (event.shiftKey) {
              if (anchor.current === null) anchor.current = visibleAt
              draftEnd.current = active.current
              setDraftLabel(`Draft range ${periods[Math.min(anchor.current, active.current)]!.name} through ${periods[Math.max(anchor.current, active.current)]!.name}. Press Enter to apply or Escape to cancel.`)
              root.current?.querySelector<HTMLButtonElement>(`[data-timeline-index="${active.current}"]`)?.focus()
            } else {
              root.current?.querySelector<HTMLButtonElement>(`[data-timeline-index="${active.current}"]`)?.focus()
            }
          } else if (event.key === 'Enter' || event.key === ' ') {
            event.preventDefault()
            if (anchor.current === null) { anchor.current = visibleAt; draftEnd.current = visibleAt; setDraftLabel(`Range starts at ${period.name}. Extend with Shift+Arrow keys, then press Enter to apply.`) }
            else if (draftEnd.current !== null) commitRange(anchor.current, draftEnd.current)
            else { anchor.current = visibleAt; draftEnd.current = visibleAt; setDraftLabel(`Range starts at ${period.name}. Extend with Shift+Arrow keys, then press Enter to apply.`) }
          } else if (event.key === 'Escape') {
            if (anchor.current !== null) { event.preventDefault(); anchor.current = null; draftEnd.current = null; setDraftLabel('Range cancelled; the committed timeline selection is unchanged.') }
            else focusGrid()
          }
        }} />
    })}
    {model.show?.timeLevel !== false && levelBox && <select ref={levelPicker} tabIndex={-1} aria-label="Timeline level" value={model.level} onChange={event => selectLevel(event.currentTarget.value as TimelineLevel)}
      onKeyDown={event => {
        if (event.key === 'Tab') { event.preventDefault(); if (event.shiftKey) lastPeriod(); else if (clearButton.current) clearButton.current.focus(); else focusGrid() }
        else if (event.key === 'Escape') { event.preventDefault(); focusGrid() }
      }}
      style={{ position: 'absolute', left: levelBox.x, top: levelBox.y, width: levelBox.width, height: levelBox.height, pointerEvents: 'auto' }}>
      {levels.map(level => <option key={level} value={level}>{level[0]!.toUpperCase() + level.slice(1)}</option>)}
    </select>}
    {model.show?.selectionLabel !== false && clearBox && <button ref={clearButton} type="button" tabIndex={-1} aria-label="Clear timeline selection" disabled={!model.selection} onClick={() => void controller.clearTimeline(model.name, new AbortController().signal)}
      onKeyDown={event => {
        if (event.key === 'Tab') { event.preventDefault(); if (event.shiftKey) { if (levelPicker.current) levelPicker.current.focus(); else lastPeriod() } else focusGrid() }
        else if (event.key === 'Escape') { event.preventDefault(); focusGrid() }
      }}
      style={{ position: 'absolute', left: clearBox.x, top: clearBox.y, width: clearBox.width, height: clearBox.height, boxSizing: 'border-box', border: '1px solid #888', borderRadius: 3, background: '#fff', color: '#242424', font: 'inherit', cursor: 'pointer', pointerEvents: 'auto' }}>Clear</button>}
    {trackBox && <button type="button" disabled tabIndex={-1} aria-label="Timeline scrolling unavailable; not measured" title="Timeline scrolling is unavailable until measured in Excel."
      style={styleAt(trackBox)} />}
    {draftLabel && <span role="status" style={{ position: 'absolute', left: 1, top: 1, width: 1, height: 1, overflow: 'hidden', clipPath: 'inset(50%)' }}>{draftLabel}</span>}
  </div>
}
