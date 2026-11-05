import { useEffect, useRef, useState, type CSSProperties, type ReactElement } from 'react'
import { logicalToViewport, type ViewportSize } from '../../../xlsx/view/scrollWindow'
import { clipBoxToPaneRegion, paneRegionClips, paneRegionKind } from '../../../xlsx/view/frozenPanes'
import type { ScrollWindow } from '../../../xlsx/view/contracts'
import { PX_PER_PT } from '../../../xlsx/pivots/index'
import { slicerLayout } from '../../../xlsx/pivots/slicers/paint'
import type { Box, SlicerModel } from '../../../xlsx/pivots/slicers/model'
import type { SlicerController } from '../../../xlsx/pivots/slicers/controller'

export interface SlicerLayerProps {
  model: SlicerModel
  /** Worksheet-logical CSS pixels, the same anchored box the canvas painter uses. */
  box: Box
  window: ScrollWindow
  viewport: ViewportSize
  controller: SlicerController
  invoker?: HTMLElement
}

const buttonStyle: CSSProperties = { position: 'absolute', boxSizing: 'border-box', margin: 0, padding: 0, overflow: 'hidden', font: 'inherit', lineHeight: 'normal', whiteSpace: 'nowrap', textOverflow: 'ellipsis', cursor: 'pointer' }

/** Accessible controls over the canvas-painted slicer. The visible item set is bounded by the drawing's viewport box. */
export function SlicerLayer({ model: savedModel, box, window, viewport, controller, invoker }: SlicerLayerProps): ReactElement | null {
  const [, tick] = useState(0)
  const root = useRef<HTMLDivElement>(null)
  useEffect(() => controller.subscribe(() => tick(value => value + 1)), [controller])
  const model = controller.model(savedModel.name) ?? savedModel
  const zoom = window.zoom
  const layout = slicerLayout(model, box, PX_PER_PT)
  const clips = paneRegionClips(window.frozen, window.logicalOrigin, zoom, viewport)
  // Worksheet-logical rectangles become viewport boxes clipped to their pane
  // region: a part scrolled under a frozen strip has no target, as on canvas.
  const place = (lx: number, ly: number, lw: number, lh: number) => {
    const at = logicalToViewport(window, { x: lx, y: ly })
    const clipped = clipBoxToPaneRegion(clips, paneRegionKind(window.frozen, lx, ly), { x: at.x, y: at.y, width: lw * zoom, height: lh * zoom })
    if (!clipped) return undefined
    if (clipped.x >= viewport.width || clipped.y >= viewport.height || clipped.x + clipped.width <= 0 || clipped.y + clipped.height <= 0) return undefined
    return clipped
  }
  const visible = layout.items.map(item => {
    const target = place(item.rect.x, item.rect.y, item.rect.width, item.rect.height)
    return target ? { item, box: target } : undefined
  }).filter((entry): entry is { item: (typeof layout.items)[number]; box: { x: number; y: number; width: number; height: number } } => entry !== undefined)
  if (!visible.length) return null
  const allIds = model.allItemIds ?? model.items.map(item => item.id)
  const selectedIds = new Set(model.selectedItemIds ?? model.items.filter(item => item.selected).map(item => item.id))
  // A restricted cache can hide behind an all-selected visible set when
  // showMissing drops no-data items: Clear stays available until everything
  // in the cache, not just everything drawn, is selected.
  const restricted = allIds.some(id => !selectedIds.has(id))
  const captionBox = model.showCaption ? place(box.x + 7 * PX_PER_PT, box.y, Math.max(0, box.width - 120), 19 * PX_PER_PT) : undefined
  const multiBox = place(Math.max(box.x, box.x + box.width - 112), box.y, 58, 19 * PX_PER_PT)
  const clearBox = place(Math.max(box.x, box.x + box.width - 54), box.y, 48, 19 * PX_PER_PT)
  const invoke = (id: number, event: { ctrlKey?: boolean; metaKey?: boolean; shiftKey?: boolean }) => {
    if (event.ctrlKey || event.metaKey || event.shiftKey) {
      controller.toggleMultiSelect(model.name, true)
      return
    }
    const selected = model.selectedItemIds ?? model.items.filter(item => item.selected).map(item => item.id)
    const allSelected = (model.allItemIds ?? []).length > 0 && selected.length === model.allItemIds!.length
    const next = allSelected ? [id] : selected.includes(id) ? selected.filter(itemId => itemId !== id) : [...selected, id]
    void controller.applySlicer(model.name, next, new AbortController().signal)
  }
  const focusOffset = (from: HTMLButtonElement, offset: number) => {
    const buttons = [...(root.current?.querySelectorAll<HTMLButtonElement>('[data-slicer-item]') ?? [])]
    if (!buttons.length) return
    const next = Math.max(0, Math.min(buttons.length - 1, buttons.indexOf(from) + offset))
    buttons[next]?.focus()
  }
  return <div ref={root} role="group" aria-label={`${model.caption ?? model.name} slicer`} data-officeview-ui="slicer-layer" data-slicer-name={model.name}
    style={{ position: 'absolute', left: 0, top: 0, width: viewport.width, height: viewport.height, overflow: 'hidden', pointerEvents: 'none' }}>
    {visible.map(({ item: placed, box: target }) => {
      const item = model.items[placed.index]!
      const paint = model.appearance?.items[placed.state]
      const name = `${item.label || '(blank)'}, ${item.selected ? 'selected' : 'not selected'}${item.noData ? ', no data' : ''}`
      return <button key={item.id} type="button" data-slicer-item="" tabIndex={-1} aria-label={name} aria-pressed={item.selected} disabled={item.noData}
        onClick={event => invoke(item.id, event)}
        onKeyDown={event => {
          if (event.key === 'ArrowDown' || event.key === 'ArrowRight') { event.preventDefault(); focusOffset(event.currentTarget, 1) }
          else if (event.key === 'ArrowUp' || event.key === 'ArrowLeft') { event.preventDefault(); focusOffset(event.currentTarget, -1) }
          else if (event.key === 'Home') { event.preventDefault(); root.current?.querySelector<HTMLButtonElement>('[data-slicer-item]')?.focus() }
          else if (event.key === 'End') { event.preventDefault(); const buttons = root.current?.querySelectorAll<HTMLButtonElement>('[data-slicer-item]'); buttons?.item(buttons.length - 1).focus() }
          else if (event.key === 'Escape') {
            event.preventDefault()
            const grid = invoker ?? root.current?.closest('[data-officeview-sheet-pane]')?.querySelector<HTMLElement>('[role="grid"]')
            grid?.focus()
          }
          else if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); invoke(item.id, event) }
        }}
        style={{ ...buttonStyle, left: target.x, top: target.y, width: target.width, height: target.height, border: `1px solid ${paint?.border ?? '#8a8886'}`, borderRadius: 3,
          background: paint?.fill ?? (item.selected ? '#d6e8f5' : '#f3f2f1'), color: paint?.text ?? '#242424', fontWeight: paint?.bold ? 700 : 400, pointerEvents: 'auto' }} />
    })}
    {visible.map(({ item: placed }) => {
      const item = model.items[placed.index]!
      const label = place(placed.text.x, placed.rect.y, Math.max(0, placed.rect.width - 12 * PX_PER_PT), placed.rect.height)
      if (!label) return null
      return <span key={`label:${item.id}`} aria-hidden="true" style={{ position: 'absolute', left: label.x,
        top: label.y, width: label.width, height: label.height,
        lineHeight: `${placed.rect.height * zoom}px`, overflow: 'hidden', whiteSpace: 'nowrap', textOverflow: 'ellipsis', pointerEvents: 'none', color: model.appearance?.items[placed.state].text ?? '#242424' }}>{item.label || '(blank)'}</span>
    })}
    {captionBox && <span aria-hidden="true" style={{ position: 'absolute', left: captionBox.x, top: captionBox.y, height: captionBox.height,
      lineHeight: `${19 * PX_PER_PT * zoom}px`, maxWidth: captionBox.width, overflow: 'hidden', whiteSpace: 'nowrap', textOverflow: 'ellipsis', pointerEvents: 'none', fontWeight: 700 }}>{model.caption ?? model.name}</span>}
    {multiBox && <button type="button" disabled aria-label="Multi-Select unavailable" aria-pressed={controller.multiSelect(model.name)} title="Not available: this interaction has not been measured in Excel."
      style={{ ...buttonStyle, left: multiBox.x, top: multiBox.y, width: multiBox.width, height: multiBox.height, pointerEvents: 'auto', background: '#f3f2f1', color: '#242424' }}>Multi-Select</button>}
    {clearBox && <button type="button" aria-label="Clear Filter" disabled={!restricted}
      onClick={() => void controller.clearSlicer(model.name, new AbortController().signal)}
      style={{ ...buttonStyle, left: clearBox.x, top: clearBox.y, width: clearBox.width, height: clearBox.height, pointerEvents: 'auto', background: '#f3f2f1', color: '#242424' }}>Clear Filter</button>}
  </div>
}
