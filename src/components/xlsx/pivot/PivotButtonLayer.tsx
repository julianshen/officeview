/**
 * Pivot buttons as controls (plan Task 28, spec §6.6, §6.7, §7.6, §8). The filter arrows, ± buttons and page-field
 * dropdowns are painted on the canvas in every mode (buttons.ts); in the interactive viewer each visible one also
 * gets a DOM hit target here: a `<button data-officeview-ui>` over its painted box, outside the grid's ARIA tree
 * (so grid selection never sees its clicks, `isViewerUiTarget`) and named by its field and action, or by the pivot's
 * refusal reason once the pivot is known not to be interactive (§7.6; the buttons stay painted).
 *
 * The layer is one roving-tabindex toolbar: Tab reaches its one stop, arrows, Home and End move between its buttons,
 * Escape returns to the grid. On a coarse pointer each target is at least 44 × 44 px, centred on its painted box.
 *
 * Lazy: this module is the pivot-controls chunk the viewer imports on a sheet with pivots; it carries the client
 * (and through it the `pivots` chunk) and the menus.
 */
import { useEffect, useMemo, useState, type CSSProperties, type KeyboardEvent, type ReactElement } from 'react'
import type { CacheIdentity } from '../../../xlsx/view/contracts'
import type { ProjectedSheetMetrics } from '../../../xlsx/view/projection'
import { logicalToViewport, type ViewportSize } from '../../../xlsx/view/scrollWindow'
import type { ScrollWindow } from '../../../xlsx/view/contracts'
import type { PivotDiagnostic, PivotTableDefinition } from '../../../xlsx/pivots/model'
import { drillBox, dropdownBox, pivotButtons, PX_PER_PT, savedPivotLayout, type ButtonPlacement } from '../../../xlsx/pivots/index'
import type { PivotClient, PivotDelta, PivotDeltaKind, PivotMenuModel } from './pivotClient'
import { PivotFieldMenu } from './PivotFieldMenu'
import { PivotPageMenu } from './PivotPageMenu'

export { createResidentPivotClient, createRemotePivotClient, unavailablePivotClient, REMOTE_PIVOTS_UNAVAILABLE } from './pivotClient'
export type { PivotClient, PivotCommit, PivotOpenResult, PivotApplyOutcome, RemotePivotPort, PivotDeltaKind } from './pivotClient'

export type PivotUiState =
  | { status: 'idle' }
  | { status: 'opening' }
  | { status: 'open'; revision: number; verified: ReadonlySet<PivotDeltaKind>; model: PivotMenuModel }
  | { status: 'refused'; diagnostic: PivotDiagnostic }
export interface PivotMenuRequest { sheet: number; pivot: number; field: number; kind: 'field' | 'page'; invoker: HTMLElement }
export interface PivotControllerHost {
  client: PivotClient
  announce(message: string): void
  openMenu(request: PivotMenuRequest): void
  closeMenu(): void
}

const IDLE: PivotUiState = { status: 'idle' }
export const PIVOT_APPLIED = 'Pivot table changed (view only; the file is unchanged).'
export const PIVOT_RESET = 'Pivot table back to its saved view.'

/**
 * Per-viewer pivot interaction state: one session per pivot, opened on its first use (spec §6.4), and the menus'
 * models. Applies run one at a time, in the order they were asked for.
 */
export class PivotController {
  private readonly states = new Map<string, PivotUiState>()
  private readonly listeners = new Set<() => void>()
  private readonly abort = new AbortController()
  private queue: Promise<unknown> = Promise.resolve()
  private disposed = false
  constructor(private readonly host: PivotControllerHost) {}
  get client(): PivotClient { return this.host.client }
  private key = (sheet: number, pivot: number) => `${sheet}:${pivot}`
  state(sheet: number, pivot: number): PivotUiState {
    const known = this.states.get(this.key(sheet, pivot))
    if (known) return known
    const refused = this.host.client.preflight?.(sheet, pivot)
    return refused ? { status: 'refused', diagnostic: refused } : IDLE
  }
  subscribe(listener: () => void): () => void { this.listeners.add(listener); return () => { this.listeners.delete(listener) } }
  private set(sheet: number, pivot: number, state: PivotUiState): void {
    if (this.disposed) return
    this.states.set(this.key(sheet, pivot), state)
    for (const listener of this.listeners) listener()
  }
  /** One notice per pivot not interactive (§7.6), deduplicated by text. */
  notices(sheet: number, count: number): string[] {
    const out = new Set<string>()
    for (let pivot = 0; pivot < count; pivot++) { const state = this.state(sheet, pivot); if (state.status === 'refused') out.add(state.diagnostic.message) }
    return [...out]
  }
  async open(sheet: number, pivot: number): Promise<PivotUiState> {
    const current = this.state(sheet, pivot)
    if (current.status === 'open' || current.status === 'refused') return current
    this.set(sheet, pivot, { status: 'opening' })
    let next: PivotUiState
    try {
      const result = await this.host.client.open(sheet, pivot, this.abort.signal)
      next = result.ok ? { status: 'open', revision: result.revision, verified: result.verifiedKinds, model: result.model } : { status: 'refused', diagnostic: result.diagnostic }
    } catch (error) {
      if (this.abort.signal.aborted) return IDLE
      next = { status: 'refused', diagnostic: { kind: 'unsupported', feature: 'pivot-session-failed', message: `This pivot table cannot be changed: ${error instanceof Error ? error.message.slice(0, 200) : 'it could not be opened'}.` } }
    }
    this.set(sheet, pivot, next)
    return next
  }
  /** A button's use: opens the pivot if needed, then applies a ± change or opens the button's menu. */
  async activate(sheet: number, pivot: number, button: ButtonPlacement, invoker: HTMLElement): Promise<void> {
    const state = await this.open(sheet, pivot)
    if (state.status !== 'open') return
    if (button.kind === 'expand' || button.kind === 'collapse') {
      if (!state.verified.has('collapse') || button.item === undefined) { this.host.announce('Expanding and collapsing this pivot table is not available.'); return }
      await this.apply(sheet, pivot, { kind: 'collapse', field: button.field, item: button.item, collapsed: button.kind === 'collapse' })
      return
    }
    this.host.openMenu({ sheet, pivot, field: button.field, kind: button.kind === 'page' ? 'page' : 'field', invoker })
  }
  /** Applies (and, through the client, commits and installs) one change; refusals are announced (S2-AJ and the rest). */
  apply(sheet: number, pivot: number, delta: PivotDelta | 'reset'): Promise<boolean> {
    const run = this.queue.then(async () => {
      if (this.disposed) return false
      const outcome = await this.host.client.apply(sheet, pivot, delta)
      if (!outcome.ok) { this.host.announce(outcome.diagnostic.message); return false }
      const state = this.state(sheet, pivot)
      if (state.status === 'open') this.set(sheet, pivot, { ...state, revision: outcome.revision, model: outcome.model })
      this.host.announce(delta === 'reset' ? PIVOT_RESET : PIVOT_APPLIED)
      return true
    }).catch(error => { this.host.announce(error instanceof Error ? error.message.slice(0, 256) : 'Pivot change failed.'); return false })
    this.queue = run
    return run
  }
  dispose(): void {
    this.disposed = true
    this.abort.abort()
    this.listeners.clear()
    this.host.client.dispose?.()
  }
}

// Placements are cached per definition object, like the painter's: a new live definition re-places its buttons.
const placed = new WeakMap<PivotTableDefinition, ButtonPlacement[]>()
function placementsOf(def: PivotTableDefinition): ButtonPlacement[] {
  let list = placed.get(def)
  if (!list) { try { list = pivotButtons(def, 'screen', savedPivotLayout(def)) } catch { list = [] } placed.set(def, list) }
  return list
}
// The ± state by cell, indexed once per definition: every rendered gridcell asks (up to 20,000 per frame).
const drills = new WeakMap<PivotTableDefinition, Map<number, boolean>>()
const CELL_KEY = 16_384
function drillsOf(def: PivotTableDefinition): Map<number, boolean> {
  let map = drills.get(def)
  if (!map) {
    map = new Map()
    for (const button of placementsOf(def)) if (button.kind === 'expand' || button.kind === 'collapse') map.set(button.row * CELL_KEY + button.col, button.kind === 'collapse')
    drills.set(def, map)
  }
  return map
}
/** aria-expanded for a pivot cell carrying a ± button (§8: expanded/collapsed state on row headers). */
export function expandedAt(defs: readonly PivotTableDefinition[] | undefined, row: number, col: number): boolean | undefined {
  for (const def of defs ?? []) { const expanded = drillsOf(def).get(row * CELL_KEY + col); if (expanded !== undefined) return expanded }
  return undefined
}

export interface PivotButtonLayerProps {
  controller: PivotController
  sheet: number
  defs: readonly PivotTableDefinition[]
  projection: ProjectedSheetMetrics
  window: ScrollWindow
  viewport: ViewportSize
  touch: boolean
  /** The cell's shown text: names a button when the client has no labels. */
  cellText?(row: number, col: number): string
  /** Escape: focus goes back to the grid. */
  onEscape(): void
}
interface Target { key: string; pivot: number; button: ButtonPlacement; box: { x: number; y: number; width: number; height: number }; action: string; refused?: string }

const TOUCH = 44
function actionName(controller: PivotController, sheet: number, pivot: number, button: ButtonPlacement, cellText: PivotButtonLayerProps['cellText']): string {
  const labels = controller.client.labels?.(sheet, pivot)
  const field = labels?.field(button.field) ?? cellText?.(button.row, button.col) ?? `Field ${button.field + 1}`
  if (button.kind === 'filter') return `Filter and sort ${field}`
  if (button.kind === 'page') return `Report filter ${field}`
  const item = (button.item !== undefined ? labels?.item(button.field, button.item) : undefined) || cellText?.(button.row, button.col) || `item ${(button.item ?? 0) + 1}`
  return `${button.kind === 'collapse' ? 'Collapse' : 'Expand'} ${item} (${field})`
}

export function PivotButtonLayer(props: PivotButtonLayerProps): ReactElement | null {
  const { controller, sheet, defs, projection, window: scroll, viewport, touch } = props
  const [, setTick] = useState(0)
  useEffect(() => controller.subscribe(() => setTick(tick => tick + 1)), [controller])
  const zoom = scroll.zoom, ppt = PX_PER_PT * zoom
  const targets: Target[] = []
  defs.forEach((def, pivot) => {
    const state = controller.state(sheet, pivot)
    for (const button of placementsOf(def)) {
      if (button.row >= projection.rows.logicalCount || button.col >= projection.cols.logicalCount) continue
      const logical = projection.sourceCellToLogical({ row: button.row, col: button.col })
      if (!logical) continue
      const at = logicalToViewport(scroll, { x: logical.x, y: logical.y })
      const cell = { x: at.x, y: at.y, width: logical.width * zoom, height: logical.height * zoom }
      if (cell.x >= viewport.width || cell.y >= viewport.height || cell.x + cell.width <= 0 || cell.y + cell.height <= 0) continue
      let box = button.kind === 'expand' || button.kind === 'collapse' ? drillBox(cell, ppt) : dropdownBox(cell, ppt)
      if (touch) {
        const width = Math.max(TOUCH, box.width), height = Math.max(TOUCH, box.height)
        box = { x: box.x + box.width / 2 - width / 2, y: box.y + box.height / 2 - height / 2, width, height }
      }
      targets.push({ key: `${pivot}:${button.row}:${button.col}`, pivot, button, box, action: actionName(controller, sheet, pivot, button, props.cellText),
        ...(state.status === 'refused' ? { refused: state.diagnostic.message } : {}) })
    }
  })
  const [focusKey, setFocusKey] = useState<string | null>(null)
  const stop = targets.find(target => target.key === focusKey) ?? targets[0]
  const layer = useMemo(() => ({ current: null as HTMLDivElement | null }), [])
  if (!targets.length) return null
  const move = (event: KeyboardEvent<HTMLButtonElement>, index: number) => {
    let next: number | undefined
    if (event.key === 'ArrowRight' || event.key === 'ArrowDown') next = Math.min(targets.length - 1, index + 1)
    else if (event.key === 'ArrowLeft' || event.key === 'ArrowUp') next = Math.max(0, index - 1)
    else if (event.key === 'Home') next = 0
    else if (event.key === 'End') next = targets.length - 1
    else if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); props.onEscape(); return }
    if (next === undefined) return
    event.preventDefault(); event.stopPropagation()
    const target = targets[next]!
    setFocusKey(target.key)
    layer.current?.querySelector<HTMLButtonElement>(`[data-pivot-key="${target.key}"]`)?.focus()
  }
  const style: CSSProperties = { position: 'absolute', left: 0, top: 0, width: viewport.width, height: viewport.height, overflow: 'hidden', pointerEvents: 'none' }
  return <div ref={node => { layer.current = node }} role="toolbar" aria-label="Pivot table controls" aria-orientation="horizontal" data-officeview-ui="pivot-layer" style={style}>
    {targets.map((target, index) => {
      const { button } = target
      const drill = button.kind === 'expand' || button.kind === 'collapse'
      const name = target.refused ?? target.action
      return <button key={target.key} type="button" data-officeview-ui="pivot-button" data-pivot-key={target.key} data-pivot={target.pivot} data-row={button.row} data-col={button.col}
        tabIndex={target === stop ? 0 : -1} aria-label={name} title={name} {...(target.refused ? { 'aria-disabled': 'true', 'aria-description': target.action } : {})}
        {...(drill ? { 'aria-expanded': button.kind === 'collapse' } : { 'aria-haspopup': 'dialog' as const })}
        onFocus={() => setFocusKey(target.key)} onKeyDown={event => move(event, index)}
        onPointerDown={event => event.stopPropagation()} onDoubleClick={event => event.stopPropagation()}
        onClick={event => { event.stopPropagation(); setFocusKey(target.key); void controller.activate(sheet, target.pivot, button, event.currentTarget) }}
        // Transparent over the painted button; the focus ring is the one visible part.
        className="ov:rounded-control ov:focus-visible:outline-2 ov:focus-visible:outline-solid ov:focus-visible:outline-offset-1 ov:focus-visible:outline-accent"
        style={{ position: 'absolute', left: target.box.x, top: target.box.y, width: target.box.width, height: target.box.height, padding: 0, margin: 0, border: 0, background: 'transparent', cursor: target.refused ? 'default' : 'pointer', pointerEvents: 'auto' }} />
    })}
  </div>
}

/** The menu a request opens, over the pivot's current model. */
export function PivotMenu({ controller, request, owner, onClose, onNotice }: { controller: PivotController; request: PivotMenuRequest; owner: CacheIdentity; onClose(): void; onNotice(message: string): void }): ReactElement | null {
  const [, setTick] = useState(0)
  useEffect(() => controller.subscribe(() => setTick(tick => tick + 1)), [controller])
  const state = controller.state(request.sheet, request.pivot)
  if (state.status !== 'open') return null
  const onApply = (delta: PivotDelta | 'reset') => { onClose(); void controller.apply(request.sheet, request.pivot, delta) }
  const common = { owner, model: state.model, verified: state.verified, field: request.field, onApply, onClose, onNotice }
  return request.kind === 'page' ? <PivotPageMenu {...common} /> : <PivotFieldMenu {...common} />
}
