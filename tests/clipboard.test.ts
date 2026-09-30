import { describe, test, expect, vi, beforeEach, afterEach } from 'vitest'
import { copyText } from '../src/core/clipboard'

const originalClipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard')
const originalExec = document.execCommand

function setClipboard(value: unknown) {
  Object.defineProperty(navigator, 'clipboard', { value, configurable: true, writable: true })
}

beforeEach(() => {
  setClipboard(undefined)
})

afterEach(() => {
  if (originalClipboard) Object.defineProperty(navigator, 'clipboard', originalClipboard)
  document.execCommand = originalExec
  vi.restoreAllMocks()
})

describe('copyText', () => {
  test('uses the async clipboard API when available', async () => {
    const writeText = vi.fn().mockResolvedValue(undefined)
    setClipboard({ writeText })
    const outcome = await copyText('hello')
    expect(writeText).toHaveBeenCalledWith('hello')
    expect(outcome).toBe('clipboard')
  })

  test('falls back to execCommand when the clipboard API rejects', async () => {
    const writeText = vi.fn().mockRejectedValue(new Error('NotAllowedError'))
    setClipboard({ writeText })
    document.execCommand = vi.fn().mockReturnValue(true) as unknown as typeof document.execCommand
    const outcome = await copyText('fallback text')
    expect(writeText).toHaveBeenCalled()
    expect(document.execCommand).toHaveBeenCalledWith('copy')
    expect(outcome).toBe('fallback')
  })

  test('reports failure instead of silently doing nothing', async () => {
    const writeText = vi.fn().mockRejectedValue(new Error('NotAllowedError'))
    setClipboard({ writeText })
    document.execCommand = vi.fn().mockReturnValue(false) as unknown as typeof document.execCommand
    const outcome = await copyText('nope')
    expect(outcome).toBe('failed')
  })

  test('reports failure when execCommand throws', async () => {
    setClipboard(undefined)
    document.execCommand = vi.fn(() => {
      throw new Error('blocked')
    }) as unknown as typeof document.execCommand
    expect(await copyText('x')).toBe('failed')
  })

  test('cleans up the fallback textarea and restores focus', async () => {
    setClipboard(undefined)
    document.execCommand = vi.fn().mockReturnValue(true) as unknown as typeof document.execCommand
    const input = document.createElement('input')
    document.body.appendChild(input)
    input.focus()
    await copyText('cleanup me')
    expect(document.querySelectorAll('textarea')).toHaveLength(0)
    expect(document.activeElement).toBe(input)
    input.remove()
  })
})