import { describe, expect, it, vi } from 'vitest'
import { SlicerController } from '../src/xlsx/pivots/slicers/controller'
import type { SlicerModel } from '../src/xlsx/pivots/slicers/model'

function model(): SlicerModel {
  return {
    name: 'Slicer_Product', cacheName: 'Slicer_Product', cacheId: 1, sourceFieldIndex: 2, revision: 0, allItemIds: [0, 1, 2, 3], selectedItemIds: [0, 1, 2, 3],
    caption: 'Product', showCaption: true, columnCount: 1, rowHeightEmu: 209550, style: 'SlicerStyleLight1', startItem: 0,
    items: [
      { id: 0, label: 'Apples', selected: true, noData: false },
      { id: 1, label: 'Bananas', selected: true, noData: false },
      { id: 2, label: 'Cherries', selected: true, noData: false },
      { id: 3, label: '(blank)', selected: true, noData: true },
    ], drives: [{ pivotTable: 'Pivot1' }],
  }
}

describe('slicer controller', () => {
  it('commits the measured click-one selection by stable item id without changing saved no-data order', async () => {
    const apply = vi.fn(async (_model: SlicerModel, _action: { kind: string; selectedItemIds?: readonly number[] }, _signal: AbortSignal) => ({ kind: 'committed' as const, revision: 1 }))
    const controller = new SlicerController([model()], { apply, announce: vi.fn() })
    expect(await controller.applySlicer('Slicer_Product', [0], new AbortController().signal)).toBe(true)
    expect(apply.mock.calls[0]?.[0].name).toBe('Slicer_Product')
    expect(apply.mock.calls[0]?.[1]).toEqual({ kind: 'selection', selectedItemIds: [0] })
    const committed = controller.model('Slicer_Product')!
    expect(committed.items.map(item => [item.id, item.selected, item.noData])).toEqual([[0, true, false], [1, false, false], [2, false, false], [3, false, true]])
    expect(committed.revision).toBe(1)
  })

  it('preserves the prior selection on stale or refused replies and refuses unmeasured repeat, modifier and multi-select actions', async () => {
    const announce = vi.fn()
    const apply = vi.fn(async () => ({ kind: 'stale' as const }))
    const controller = new SlicerController([model()], { apply, announce })
    expect(await controller.applySlicer('Slicer_Product', [0], new AbortController().signal)).toBe(false)
    expect(controller.model('Slicer_Product')!.items.every(item => item.selected)).toBe(true)
    expect(await controller.applySlicer('Slicer_Product', [0, 1], new AbortController().signal)).toBe(false)
    expect(controller.toggleMultiSelect('Slicer_Product', true)).toBe(false)
    expect(controller.multiSelect('Slicer_Product')).toBe(false)
    expect(apply).toHaveBeenCalledTimes(1)
    expect(announce).toHaveBeenCalled()
  })

  it('clears the committed selection and exposes filtered state to its linked cache field', async () => {
    const apply = vi.fn(async (_model: SlicerModel, action: { kind: string }) => ({ kind: 'committed' as const, revision: action.kind === 'clear' ? 2 : 1 }))
    const controller = new SlicerController([model()], { apply, announce: vi.fn() })
    await controller.applySlicer('Slicer_Product', [0], new AbortController().signal)
    expect(controller.isFieldFiltered(1, 2)).toBe(true)
    expect(await controller.clearSlicer('Slicer_Product', new AbortController().signal)).toBe(true)
    expect(controller.model('Slicer_Product')!.items.every(item => item.selected)).toBe(true)
    expect(controller.isFieldFiltered(1, 2)).toBe(false)
  })
})
