/** 1d Task 6 (§16 T2): a byte source never touches the stream path. Pixel
 * identity against the pre-streaming baseline is scripts/stream-s1/local-identity.ts. */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { cleanup, render, waitFor } from '@testing-library/react'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { OfficeFile } from '../../src/components/OfficeFile'

const h = vi.hoisted(() => ({ imported: 0 }))
vi.mock('../../src/stream/remote/open', () => { h.imported++; return { openRemoteDocument: () => { throw new Error('must not open') } } })

afterEach(() => { cleanup(); vi.unstubAllGlobals() })
const bytes = (name: string) => new Uint8Array(readFileSync(resolve(__dirname, '../../corpus', name)))

describe('local loads stay local', () => {
  for (const name of ['poi-55733.docx', 'lo-calc-calcfields.xlsx', 'pypptx-test.pptx']) it(`${name}: no socket, no remote chunk, no stream notice`, async () => {
    const ws = vi.fn()
    vi.stubGlobal('WebSocket', ws)
    const view = render(<OfficeFile data={bytes(name)} />)
    await waitFor(() => expect(view.container.querySelector('[data-officeview-root]')).toBeTruthy(), { timeout: 15_000 })
    expect(ws).not.toHaveBeenCalled()
    expect(h.imported).toBe(0)
    expect(view.container.textContent).not.toContain('text unavailable in this viewing mode')
  }, 20_000)
})
