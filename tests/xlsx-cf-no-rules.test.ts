import { expect, it, vi } from 'vitest'
import { prepareCfPaint } from '../src/xlsx/cf/render'
import { createCfMemory } from '../src/xlsx/cf/limits'
it('100 no-CF preparations do zero source reads and reserve zero bytes', async () => {
  const memory = createCfMemory(), reserve = vi.spyOn(memory, 'reserve')
  const store = { get stamp() { throw Error('no-CF stamp read') }, readRows: vi.fn(() => { throw Error('no-CF read') }) }
  for (let i = 0; i < 100; i++) expect(await prepareCfPaint({ name: 'plain', rows: [], cols: [], merges: [], mergeRanges: [] }, { memory, store: store as never })).toBeUndefined()
  expect(reserve).not.toHaveBeenCalled(); expect(store.readRows).not.toHaveBeenCalled()
})

it('100 ordinary mounted frames never load or prepare CF', async () => {
  const { createElement } = await import('react')
  const { render, waitFor } = await import('@testing-library/react')
  const { default: Viewer } = await import('../src/components/xlsx/InteractiveSheetDoc')
  const loader = await import('../src/xlsx/cfLoader')
  const prepare = vi.spyOn(loader, 'ensureConditionalFormatting')
  const doc = { sheets: [{ name: 'Plain', rows: [{ index: 0, cells: [{ row: 0, col: 0, ref: 'A1', styleIndex: 0, value: 1 }] }], cols: [], merges: [], mergeRanges: [] }] }
  const frames = vi.fn()
  const ui = render(createElement(Viewer, { doc, onFrame: frames }))
  try {
    await waitFor(() => expect(frames).toHaveBeenCalled())
    for (let i = 1; i <= 100; i++) {
      const before = frames.mock.calls.length
      ui.rerender(createElement(Viewer, { doc, onFrame: frames, views: { '1': { revision: i } } }))
      await waitFor(() => expect(frames.mock.calls.length).toBeGreaterThan(before))
    }
    expect(prepare).not.toHaveBeenCalled()
  } finally { ui.unmount(); prepare.mockRestore() }
}, 30_000)
