/** Stage 3b Task 1 — print policy (page caps) and its carriage in the manifest. */
import { afterAll, describe, expect, test } from 'bun:test'
import { printPolicy, effectivePolicy } from '../src/hooks'
import { Client, disposePool, start } from './harness'

afterAll(() => disposePool())

describe('printPolicy (fail closed, like copyPolicy)', () => {
  test('undefined and true are uncapped', () => {
    expect(printPolicy(undefined)).toEqual({ maxPages: null, maxSessionPages: null })
    expect(printPolicy(true)).toEqual({ maxPages: null, maxSessionPages: null })
  })
  test('false is off', () => { expect(printPolicy(false)).toBe(false) })
  test('caps resolve', () => {
    expect(printPolicy({ maxPages: 3 })).toEqual({ maxPages: 3, maxSessionPages: null })
    expect(printPolicy({ maxPages: 3, maxSessionPages: 5 })).toEqual({ maxPages: 3, maxSessionPages: 5 })
  })
  test.each([
    [{ maxPages: 0 }], [{ maxPages: 1.5 }], [{ maxPages: -1 }], [{ maxPages: 3, maxSessionPages: 0 }],
    [{ maxPages: 3, maxSessionPages: 'x' }], [{}], ['x'], [null], [[]], [1],
  ])('%j fails closed to off', (v) => { expect(printPolicy(v)).toBe(false) })
  test('effectivePolicy carries the resolved shape', () => {
    expect(effectivePolicy({ policy: { print: { maxPages: 2 } } }).print).toEqual({ maxPages: 2, maxSessionPages: null })
    expect(effectivePolicy({ policy: { print: 'x' as never } }).print).toBe(false)
  })
})

describe('the manifest carries the print policy', () => {
  test('capped print', async () => {
    const h = await start({ policy: { print: { maxPages: 2, maxSessionPages: 4 } } }), c = new Client(h.url)
    const m = await c.hello()
    expect(m.can.print).toBe(true)
    expect(m.caps.print).toEqual({ maxPages: 2, maxSessionPages: 4 })
    c.close(); h.stop()
  })
  test('default print is uncapped', async () => {
    const h = await start(), c = new Client(h.url)
    const m = await c.hello()
    expect(m.can.print).toBe(true)
    expect(m.caps.print).toEqual({ maxPages: null, maxSessionPages: null })
    c.close(); h.stop()
  })
  test('print: false', async () => {
    const h = await start({ policy: { print: false } }), c = new Client(h.url)
    const m = await c.hello()
    expect(m.can.print).toBe(false)
    expect(m.caps.print).toBeNull()
    c.close(); h.stop()
  })
})
