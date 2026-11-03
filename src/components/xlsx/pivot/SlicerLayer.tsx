import { useEffect, useRef, useState, type CSSProperties, type ReactElement } from 'react'
import { logicalToViewport, type ViewportSize } from '../../../xlsx/view/scrollWindow'
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
  const origin = logicalToViewport(window, { x: box.x, y: box.y })
  const viewBox = { x: origin.x, y: origin.y, width: box.width * zoom, height: box.height * zoom }
  const layout = slicerLayout(model, box, PX_PER_PT * zoom)
  const visible = layout.items.map(item => {
    const at = logicalToViewport(window, { x: item.rect.x, y: item.rect.y })
    return { item, box: { x: at.x, y: at.y, width: item.rect.width * zoom, height: item.rect.height * zoom } }
  }).filter(entry => entry.box.x + entry.box.width > 0 && entry.box.y + entry.box.height > 0 && entry.box.x < viewport.width && entry.box.y < viewport.height)
  if (!visible.length) return null
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
      return <span key={`label:${item.id}`} aria-hidden="true" style={{ position: 'absolute', left: logicalToViewport(window, { x: placed.text.x, y: placed.rect.y }).x,
        top: logicalToViewport(window, { x: placed.rect.x, y: placed.rect.y }).y, width: Math.max(0, placed.rect.width - 12 * PX_PER_PT * zoom), height: placed.rect.height * zoom,
        lineHeight: `${placed.rect.height * zoom}px`, overflow: 'hidden', whiteSpace: 'nowrap', textOverflow: 'ellipsis', pointerEvents: 'none', color: model.appearance?.items[placed.state].text ?? '#242424' }}>{item.label || '(blank)'}</span>
    })}
    {model.showCaption && <span aria-hidden="true" style={{ position: 'absolute', left: viewBox.x + 7 * PX_PER_PT * zoom, top: viewBox.y, height: 19 * PX_PER_PT * zoom,
      lineHeight: `${19 * PX_PER_PT * zoom}px`, maxWidth: Math.max(0, viewBox.width - 120 * zoom), overflow: 'hidden', whiteSpace: 'nowrap', textOverflow: 'ellipsis', pointerEvents: 'none', fontWeight: 700 }}>{model.caption ?? model.name}</span>}
    <button type="button" disabled aria-label="Multi-Select unavailable" aria-pressed={controller.multiSelect(model.name)} title="Not available: this interaction has not been measured in Excel."
      style={{ ...buttonStyle, left: Math.max(viewBox.x, viewBox.x + viewBox.width - 112 * zoom), top: viewBox.y, width: 58 * zoom, height: 19 * PX_PER_PT * zoom, pointerEvents: 'auto', background: '#f3f2f1', color: '#242424' }}>Multi-Select</button>
    <button type="button" aria-label="Clear Filter" disabled={model.items.every(item => item.selected)}
      onClick={() => void controller.clearSlicer(model.name, new AbortController().signal)}
      style={{ ...buttonStyle, left: Math.max(viewBox.x, viewBox.x + viewBox.width - 54 * zoom), top: viewBox.y, width: 48 * zoom, height: 19 * PX_PER_PT * zoom, pointerEvents: 'auto', background: '#f3f2f1', color: '#242424' }}>Clear Filter</button>
  </div>
}
