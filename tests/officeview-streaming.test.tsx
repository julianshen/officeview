import { describe, test, expect, vi, afterEach } from 'vitest'
import { render, waitFor, act, fireEvent } from '@testing-library/react'
import { OfficeFile } from '../src/components/OfficeFile'
import { PROTECTION_HEADER } from '../src/core/stream'
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

describe('<OfficeFile> server-driven protection', () => {
  afterEach(() => {
    vi.unstubAllGlobals()
  })

  function stubDownload(bytes: Uint8Array, protection?: string): void {
    vi.stubGlobal('fetch', async (): Promise<Response> => {
      const headers: Record<string, string> = { 'content-length': String(bytes.length) }
      if (protection !== undefined) headers[PROTECTION_HEADER] = protection
      return new Response(chunkedStream(bytes, 128), { headers })
    })
  }

  const rootOf = (container: HTMLElement): HTMLElement => {
    const root = container.querySelector('[data-officeview-root]')
    expect(root).not.toBeNull()
    return root as HTMLElement
  }

  test('a no-copy, no-print header disables copy and print even with permissive props', async () => {
    const bytes = await buildDocx([{ runs: [{ text: 'guarded content here' }] }])
    stubDownload(bytes, 'no-copy, no-print')
    const { container } = render(
      <OfficeFile data={{ url: 'https://files.example.com/guarded.docx' }} allowCopy={true} allowPrint={true} />,
    )
    await waitFor(() => expect(container.querySelectorAll('canvas').length).toBeGreaterThan(0), { timeout: 30_000 })
    // copy is off despite the explicit prop: the server denial wins, so there
    // is no prop for devtools to flip back on
    expect(rootOf(container).style.userSelect).toBe('none')
    // print is off: the print guard hides the viewer from printouts…
    expect(document.head.querySelector('[data-officeview-print-guard]')).not.toBeNull()
    // …and the print shortcut is swallowed
    expect(fireEvent.keyDown(rootOf(container), { key: 'p', ctrlKey: true })).toBe(false)
  })

  test('props alone still disable copy/print when the server says nothing', async () => {
    const bytes = await buildDocx([{ runs: [{ text: 'prop-guarded content' }] }])
    stubDownload(bytes)
    const { container } = render(<OfficeFile data={{ url: 'https://files.example.com/open.docx' }} allowCopy={false} allowPrint={false} />)
    await waitFor(() => expect(container.querySelectorAll('canvas').length).toBeGreaterThan(0), { timeout: 30_000 })
    expect(rootOf(container).style.userSelect).toBe('none')
    expect(document.head.querySelector('[data-officeview-print-guard]')).not.toBeNull()
  })

  test('switching to an unprotected source lifts the server policy', async () => {
    const guarded = await buildDocx([{ runs: [{ text: 'guarded first' }] }])
    const open = await buildDocx([{ runs: [{ text: 'open second' }] }])
    stubDownload(guarded, 'no-copy')
    const { container, rerender } = render(
      <OfficeFile data={{ url: 'https://files.example.com/guarded.docx' }} />,
    )
    await waitFor(() => expect(container.querySelectorAll('canvas').length).toBeGreaterThan(0), { timeout: 30_000 })
    expect(rootOf(container).style.userSelect).toBe('none')
    // a stale fetch resolving late must not re-apply the old policy either:
    // the stored policy is tagged with its source
    stubDownload(open)
    rerender(<OfficeFile data={open} />)
    await waitFor(() => expect(rootOf(container).style.userSelect).not.toBe('none'), { timeout: 30_000 })
    expect(document.head.querySelector('[data-officeview-print-guard]')).toBeNull()
  })

  test('a stale permissive download resolving late cannot re-enable copy/print', async () => {
    // regression: protection events from a cancelled load were forwarded
    // through the latest callback, which tagged the old permissive policy
    // with the current data — re-enabling copy/print on the live document
    const permissive = await buildDocx([{ runs: [{ text: 'permissive first' }] }])
    const guarded = await buildDocx([{ runs: [{ text: 'guarded second' }] }])
    let resolveSlow!: (response: Response) => void
    const slowGate = new Promise<Response>((resolve) => {
      resolveSlow = resolve
    })
    vi.stubGlobal('fetch', (input: unknown): Promise<Response> => {
      if (String(input).includes('slow-permissive')) return slowGate
      return Promise.resolve(
        new Response(chunkedStream(guarded, 64), {
          headers: { 'content-length': String(guarded.length), [PROTECTION_HEADER]: 'no-copy, no-print' },
        }),
      )
    })
    const { container, rerender } = render(
      <OfficeFile data={{ url: 'https://files.example.com/slow-permissive.docx' }} />,
    )
    // switch before the slow download resolves; the guarded doc loads first
    rerender(<OfficeFile data={{ url: 'https://files.example.com/guarded.docx' }} />)
    await waitFor(() => expect(container.querySelectorAll('canvas').length).toBeGreaterThan(0), { timeout: 30_000 })
    expect(rootOf(container).style.userSelect).toBe('none')
    expect(document.head.querySelector('[data-officeview-print-guard]')).not.toBeNull()
    // now let the old permissive download finish: the live policy must not budge
    await act(async () => {
      resolveSlow(
        new Response(chunkedStream(permissive, 64), { headers: { 'content-length': String(permissive.length) } }),
      )
      await new Promise((r) => setTimeout(r, 100))
    })
    expect(rootOf(container).style.userSelect).toBe('none')
    expect(document.head.querySelector('[data-officeview-print-guard]')).not.toBeNull()
    expect(fireEvent.keyDown(rootOf(container), { key: 'p', ctrlKey: true })).toBe(false)
  })
})
