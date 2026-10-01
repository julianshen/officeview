import { describe, test, expect } from 'vitest'
import { render, waitFor, act } from '@testing-library/react'
import { OfficeFile } from '../src/components/OfficeFile'
import { buildDocx } from '../src/testdata/ooxml-builders'

function chunkedStream(bytes: Uint8Array, chunkSize: number, delayMs = 0): ReadableStream<Uint8Array> {
  let at = 0
  return new ReadableStream<Uint8Array>({
    async pull(controller) {
      if (delayMs > 0) await new Promise((r) => setTimeout(r, delayMs))
      if (at >= bytes.length) {
        controller.close()
        return
      }
      controller.enqueue(bytes.subarray(at, at + chunkSize))
      at += chunkSize
    },
  })
}

describe('<OfficeFile> streaming', () => {
  test('the function form of loading receives live progress (deterministic)', async () => {
    const bytes = await buildDocx([{ runs: [{ text: 'progressive' }] }])
    let controller!: ReadableStreamDefaultController<Uint8Array>
    const stream = new ReadableStream<Uint8Array>({ start(c) { controller = c } })
    const { container } = render(
      <OfficeFile data={stream} loading={({ loaded }) => <div data-testid="progress">{loaded}</div>} />,
    )
    // bare stream: no total is knowable, loaded starts at 0
    expect(container.querySelector('[data-testid="progress"]')?.textContent).toBe('0')
    // each pushed chunk updates the rendered count while still loading
    await act(async () => { controller.enqueue(bytes.subarray(0, 100)) })
    expect(container.querySelector('[data-testid="progress"]')?.textContent).toBe('100')
    expect(container.querySelectorAll('canvas')).toHaveLength(0)
    // closing completes the load: real render replaces the progress element
    await act(async () => { controller.enqueue(bytes.subarray(100)); controller.close() })
    await waitFor(() => expect(container.querySelectorAll('canvas').length).toBeGreaterThan(0), { timeout: 30_000 })
    expect(container.querySelector('[data-testid="progress"]')).toBeNull()
  })

  test('onProgress fires with increasing loaded counts', async () => {
    const bytes = await buildDocx([{ runs: [{ text: 'counting' }] }])
    const counts: number[] = []
    render(<OfficeFile data={chunkedStream(bytes, 256)} loading={null} />)
    const el = document.querySelector('[data-testid="progress"]')
    void el
    // observe through the exported hook path instead: loadOfficeFile onProgress
    const { loadOfficeFile } = await import('../src/components/OfficeFile')
    await loadOfficeFile(chunkedStream(bytes, 256), { onProgress: (p) => counts.push(p.loaded) })
    expect(counts.length).toBeGreaterThan(1)
    for (let i = 1; i < counts.length; i++) expect(counts[i]).toBeGreaterThanOrEqual(counts[i - 1])
    expect(counts[counts.length - 1]).toBe(bytes.length)
  })

  test('static loading element still works (backward compatible)', async () => {
    const bytes = await buildDocx([{ runs: [{ text: 'static' }] }])
    let controller!: ReadableStreamDefaultController<Uint8Array>
    const stream = new ReadableStream<Uint8Array>({ start(c) { controller = c } })
    const { container } = render(<OfficeFile data={stream} loading={<div>spinner</div>} />)
    expect(container.textContent).toContain('spinner')
    await act(async () => { controller.enqueue(bytes); controller.close() })
    await waitFor(() => expect(container.querySelectorAll('canvas').length).toBeGreaterThan(0), { timeout: 30_000 })
    expect(container.textContent).not.toContain('spinner')
  })

  test('an existing ArrayBuffer source is unchanged', async () => {
    const bytes = await buildDocx([{ runs: [{ text: 'bytes path' }] }])
    const { container } = render(<OfficeFile data={bytes.slice().buffer} />)
    await waitFor(() => expect(container.querySelectorAll('canvas').length).toBeGreaterThan(0))
  })
})
