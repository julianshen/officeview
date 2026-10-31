/**
 * Pivot buttons as controls (plan Task 28, spec §6.6, §6.7, §7.6, §8). The filter arrows, ± buttons and page-field
 * dropdowns are painted on the canvas in every mode (buttons.ts); in the interactive viewer each visible one also
 * gets a DOM hit target here: a `<button data-officeview-ui>` over its painted box, outside the grid's ARIA tree
 * (so grid selection never sees its clicks, `isViewerUiTarget`) and named by its field and action, or by the pivot's
 * refusal reason once the pivot is known not to be interactive (§7.6; the buttons stay painted).
 *
 * Ruling S3-J: the grid keeps ONE Tab stop. The buttons are for the pointer only (tabIndex -1, still named); the
 * keyboard reaches each control from its gridcell: Alt+↓ opens the cell's filter or page menu, Alt+→ expands and
 * Alt+← collapses its item (`controlsAt`), and the gridcell's accessible description announces the control
 * (`cellControlDescription`). On a coarse pointer a target is 44 px wide and as tall as 44 px within its own row band,
 * so a ± on the next row never swallows this one.
 *
 * Lazy: this module is the pivot-controls chunk the viewer imports on a sheet with pivots; it carries the client
 * (and through it the `pivots` chunk) and the menus.
 */
import { useEffect, useState, type CSSProperties, type ReactElement } from 'react'
import type { CacheIdentity } from '../../../xlsx/view/contracts'
import type { ProjectedSheetMetrics } from '../../../xlsx/view/projection'
import { logicalToViewport, type ViewportSize } from '../../../xlsx/view/scrollWindow'
import type { ScrollWindow } from '../../../xlsx/view/contracts'
import type { PivotDiagnostic, PivotTableDefinition } from '../../../xlsx/pivots/model'
import { drillBox, dropdownBox, pivotButtons, PX_PER_PT, savedPivotLayout, type ButtonPlacement } from '../../../xlsx/pivots/index'
import type { PivotClient, PivotDelta, PivotDeltaKind, PivotMenuModel } from './pivotClient'
import { PivotFieldMenu } from './PivotFieldMenu'
import { PivotPageMenu } from './PivotPageMenu'

export { createResidentPivotClient, createRemotePivotClient, remotePivotPort, unavailablePivotClient, REMOTE_PIVOTS_UNAVAILABLE } from './pivotClient'
export type { PivotClient, PivotCommit, PivotOpenResult, PivotApplyOutcome, RemotePivotPort, PivotDeltaKind } from './pivotClient'

export type PivotUiState =
  | { status: 'idle' }
  | { status: 'opening' }
  | { status: 'open'; revision: number; verified: ReadonlySet<PivotDeltaKind>; model: PivotMenuModel }
  | { status: 'refused'; diagnostic: PivotDiagnostic }
/** `invoker` gets focus back (the grid, for a keyboard route); `anchor` places the menu (the pointer button). */
export interface PivotMenuRequest { sheet: number; pivot: number; field: number; kind: 'field' | 'page'; invoker: HTMLElement; anchor?: HTMLElement }
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
  async activate(sheet: number, pivot: number, button: ButtonPlacement, invoker: HTMLElement, anchor?: HTMLElement): Promise<void> {
    const state = await this.open(sheet, pivot)
    if (state.status !== 'open') return
    if (button.kind === 'expand' || button.kind === 'collapse') {
      if (!state.verified.has('collapse') || button.item === undefined) { this.host.announce('Expanding and collapsing this pivot table is not available.'); return }
      await this.apply(sheet, pivot, { kind: 'collapse', field: button.field, item: button.item, collapsed: button.kind === 'collapse' })
      return
    }
    this.host.openMenu({ sheet, pivot, field: button.field, kind: button.kind === 'page' ? 'page' : 'field', invoker, ...(anchor ? { anchor } : {}) })
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
// Every control by cell, per definition: the keyboard route and the gridcell descriptions ask per cell.
const controls = new WeakMap<PivotTableDefinition, Map<number, ButtonPlacement[]>>()
function controlsOf(def: PivotTableDefinition): Map<number, ButtonPlacement[]> {
  let map = controls.get(def)
  if (!map) {
    map = new Map()
    for (const button of placementsOf(def)) { const key = button.row * CELL_KEY + button.col; const list = map.get(key); if (list) list.push(button); else map.set(key, [button]) }
    controls.set(def, map)
  }
  return map
}
/** The pivot controls a cell carries (usually one), with their pivot index. */
export function controlsAt(defs: readonly PivotTableDefinition[] | undefined, row: number, col: number): Array<{ pivot: number; button: ButtonPlacement }> {
  const out: Array<{ pivot: number; button: ButtonPlacement }> = []
  defs?.forEach((def, pivot) => { for (const button of controlsOf(def).get(row * CELL_KEY + col) ?? []) out.push({ pivot, button }) })
  return out
}
/** What a gridcell with a pivot control announces in its accessible description (Ruling S3-J). */
export function cellControlDescription(controller: PivotController, sheet: number, defs: readonly PivotTableDefinition[] | undefined, row: number, col: number, cellText?: (row: number, col: number) => string): string | undefined {
  const found = controlsAt(defs, row, col)
  if (!found.length) return undefined
  const parts = found.map(({ pivot, button }) => {
    const labels = controller.client.labels?.(sheet, pivot)
    const field = labels?.field(button.field) ?? cellText?.(row, col) ?? `Field ${button.field + 1}`
    return button.kind === 'filter' ? `Filter: ${field}, Alt+Down to open` : button.kind === 'page' ? `Report filter: ${field}, Alt+Down to open`
      : button.kind === 'collapse' ? 'Expanded, Alt+Left to collapse' : 'Collapsed, Alt+Right to expand'
  })
  const state = controller.state(sheet, found[0]!.pivot)
  return state.status === 'refused' ? `${parts.join('; ')}; not available: ${state.diagnostic.message}` : parts.join('; ')
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
}
interface Target { key: string; pivot: number; button: ButtonPlacement; box: { x: number; y: number; width: number; height: number }; action: string; refused?: string }

const TOUCH = 44
function actionName(controller: PivotController, sheet: number, pivot: number, button: ButtonPlacement, cellText: PivotButtonLayerProps['cellText']): string {
  const labels = controller.client.labels?.(sheet, pivot)
  const field = labels?.field(button.field) ?? cellText?.(button.row, button.col) ?? `Field ${button.field + 1}`
  if (button.kind === 'filter') return `Filter and sort ${field}`
  if (button.kind === 'page') return `Report filter ${field}`
  const item = (button.item !== undefined ? labels?.item(button.field, button.item) : undefined) || cellText?.(button.row, button.col) || `item ${(button.item ?? 0) + 1}`
  // Without labels (a worker session not opened yet) the field's name is unknown here: the item's text alone.
  return `${button.kind === 'collapse' ? 'Collapse' : 'Expand'} ${item}${labels ? ` (${field})` : ''}`
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
        // 44 px wide; as tall as 44 px within the cell's own row band, so the next row's target never covers it.
        const width = Math.max(TOUCH, box.width)
        const top = Math.max(cell.y, box.y + box.height / 2 - TOUCH / 2), bottom = Math.min(cell.y + cell.height, box.y + box.height / 2 + TOUCH / 2)
        box = { x: box.x + box.width / 2 - width / 2, y: Math.min(top, box.y), width, height: Math.max(bottom, box.y + box.height) - Math.min(top, box.y) }
      }
      targets.push({ key: `${pivot}:${button.row}:${button.col}`, pivot, button, box, action: actionName(controller, sheet, pivot, button, props.cellText),
        ...(state.status === 'refused' ? { refused: state.diagnostic.message } : {}) })
    }
  })
  if (!targets.length) return null
  const style: CSSProperties = { position: 'absolute', left: 0, top: 0, width: viewport.width, height: viewport.height, overflow: 'hidden', pointerEvents: 'none' }
  return <div data-officeview-ui="pivot-layer" style={style}>
    {targets.map(target => {
      const { button } = target
      const drill = button.kind === 'expand' || button.kind === 'collapse'
      const name = target.refused ?? target.action
      // Pointer only (Ruling S3-J): never a Tab stop; the keyboard uses the gridcell's Alt+arrow route.
      return <button key={target.key} type="button" data-officeview-ui="pivot-button" data-pivot-key={target.key} data-pivot={target.pivot} data-row={button.row} data-col={button.col}
        tabIndex={-1} aria-label={name} title={name} {...(target.refused ? { 'aria-disabled': 'true', 'aria-description': target.action } : {})}
        {...(drill ? { 'aria-expanded': button.kind === 'collapse' } : { 'aria-haspopup': 'dialog' as const })}
        onPointerDown={event => event.stopPropagation()} onDoubleClick={event => event.stopPropagation()}
        onClick={event => { event.stopPropagation(); void controller.activate(sheet, target.pivot, button, event.currentTarget) }}
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
