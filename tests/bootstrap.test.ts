import { describe, test, expect } from 'vitest'

describe('officeview bootstrap', () => {
  test('environment sanity', () => {
    expect(document).toBeDefined()
    expect(window).toBeDefined()
  })

  test('canvas 2d context is available', () => {
    const canvas = document.createElement('canvas')
    expect(canvas.getContext('2d')).toBeTruthy()
  })
})
