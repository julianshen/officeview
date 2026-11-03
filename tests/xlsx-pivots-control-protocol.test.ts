import { describe, expect, it } from 'vitest'
import { VIEWER_COMMAND_KINDS, isViewerCommand, isViewerReply } from '../src/worker/viewerProtocol'
import type { WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'

const controlApply = (overrides: Record<string, unknown> = {}) => ({
  kind: 'controlApply', requestId: 'control-1', docId: 'doc-1', revision: 1, grantBytes: 1024,
  selection: { kind: 'slicerSelection', name: 'Slicer_Product', selectedItemIds: [0, 2] },
  targets: [{ sheetId: 'sheet-1', pivot: 0, name: 'Sales', cacheId: 7, baseRevision: 0, requestRevision: 1, delta: { kind: 'items', field: 0, hidden: [1] } }],
  ...overrides,
})

describe('linked-control worker protocol', () => {
  it('registers the stage/commit/cancel commands and accepts a bounded request', () => {
    expect(VIEWER_COMMAND_KINDS).toEqual(expect.arrayContaining(['controlApply', 'controlCommit', 'controlCancel']))
    expect(isViewerCommand(controlApply())).toBe(true)
    expect(isViewerCommand({ kind: 'controlCommit', requestId: 'control-2', docId: 'doc-1', revision: 1 })).toBe(true)
    expect(isViewerCommand({ kind: 'controlCancel', requestId: 'control-3', docId: 'doc-1', revision: 1 })).toBe(true)
    expect(isViewerCommand(controlApply({ selection: { kind: 'slicerClear', name: 'Slicer_Product' } }))).toBe(true)
    expect(isViewerCommand(controlApply({ selection: { kind: 'timelineClear', name: 'NativeTimeline_Date' } }))).toBe(true)
  })

  it('rejects over-limit, duplicate or malformed fan-out before the worker sees it', () => {
    expect(isViewerCommand(controlApply({ targets: Array.from({ length: 9 }, (_, pivot) => ({ sheetId: 's', pivot, name: `P${pivot}`, cacheId: 7, baseRevision: 0, requestRevision: 1, delta: 'reset' })) }))).toBe(false)
    expect(isViewerCommand(controlApply({ targets: [
      { sheetId: 's', pivot: 0, name: 'P', cacheId: 7, baseRevision: 0, requestRevision: 1, delta: 'reset' },
      { sheetId: 's', pivot: 0, name: 'P', cacheId: 7, baseRevision: 0, requestRevision: 1, delta: 'reset' },
    ] }))).toBe(false)
    expect(isViewerCommand(controlApply({ selection: { kind: 'slicerSelection', name: 'S', selectedItemIds: Array.from({ length: 10_001 }, (_, i) => i) } }))).toBe(false)
    expect(isViewerCommand(controlApply({ selection: { kind: 'timelineSelection', name: 'T', start: 20, end: 10 } }))).toBe(false)
    expect(isViewerCommand(controlApply({ selection: { kind: 'slicerSelection', name: 'S', selectedItemIds: [0], modifier: true } }))).toBe(false)
    expect(isViewerCommand(controlApply({ selection: { kind: 'timelineSelection', name: 'T', start: 1, end: 2, action: 'drag' } }))).toBe(false)
    expect(isViewerCommand(controlApply({ grantBytes: 0 }))).toBe(false)
    expect(isViewerCommand({ kind: 'controlCommit', requestId: 'control-2', docId: 'doc-1', revision: 0 })).toBe(false)
  })

  it('keeps control outcomes correlated to document owner and command kind', () => {
    const result: WorkerReply<'controlApply'> = {
      kind: 'complete', commandKind: 'controlApply', requestId: 'control-1', owner: { docId: 'doc-1' }, diagnostics: [],
      result: { kind: 'controlApply', revision: 1, selection: { kind: 'slicerSelection', name: 'Slicer_Product', selectedItemIds: [0] }, targets: [] },
    }
    expect(isViewerReply(result)).toBe(true)
    expect(isViewerReply({ ...result, commandKind: 'controlCommit' })).toBe(true) // command-kind matching is enforced by ViewerClient.
    const command: Extract<WorkerCommand, { kind: 'controlApply' }> = { ...controlApply(), kind: 'controlApply' } as Extract<WorkerCommand, { kind: 'controlApply' }>
    expect(command.docId).toBe('doc-1')
  })
})
