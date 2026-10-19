/**
 * Final review M4: useMotionRef reaches into motion's undocumented
 * `visualElementStore`. Pin its shape so a motion bump that drops or changes
 * it is caught here, and prove the ref is guarded if it is ever missing.
 */
import { describe, expect, it, vi } from 'vitest'
import { render } from '@testing-library/react'
import * as motion from 'motion/react'

describe('motion import shape', () => {
  it('motion/react exports visualElementStore with a Map-like delete', () => {
    expect(motion.visualElementStore).toBeDefined()
    expect(typeof motion.visualElementStore.delete).toBe('function')
    expect(typeof motion.visualElementStore.get).toBe('function')
  })

  it('useMotionRef unmounts cleanly when visualElementStore is absent', async () => {
    vi.resetModules()
    vi.doMock('motion/react', async (importOriginal) => ({ ...await importOriginal<typeof import('motion/react')>(), visualElementStore: undefined }))
    try {
      const { useMotionRef } = await import('../src/components/motion/presets')
      function Probe() { const ref = useMotionRef<HTMLDivElement>(); return <div ref={ref} /> }
      const ui = render(<Probe />)
      expect(() => ui.unmount()).not.toThrow()
    } finally {
      vi.doUnmock('motion/react')
      vi.resetModules()
    }
  })
})
