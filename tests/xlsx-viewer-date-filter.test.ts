import { describe, expect, it } from 'vitest'
import { lowerDateCondition, matchesDateInterval } from '../src/xlsx/view/datePredicates'
import { encodeCivil } from '../src/xlsx/formula/serial'
import type { CalendarDay, DatePreset, TypedCondition } from '../src/xlsx/view/contracts'
const day = (year: number, month: number, day: number): CalendarDay => ({ year, month, day })
const serial = (date: CalendarDay, dateSystem: '1900' | '1904' = '1900') => encodeCivil(date, dateSystem) as number
const ctx = { dateSystem: '1900' as const, timeZone: 'America/New_York', nowMs: Date.parse('2026-03-08T16:00:00Z') }
const lower = (condition: Extract<TypedCondition, { kind: 'date' | 'datePreset' }>) => lowerDateCondition(condition, ctx)
describe('date filter calendar intervals', () => {
  it('includes midday today and re-resolves the injected clock', () => {
    const today = lower({ kind: 'datePreset', preset: 'today' })
    expect(today).toEqual({ start: serial(day(2026, 3, 8)), end: serial(day(2026, 3, 9)), complement: false })
    expect(matchesDateInterval(today.start + 0.5, today)).toBe(true)
    expect(matchesDateInterval(today.end, today)).toBe(false)
    expect(lowerDateCondition({ kind: 'datePreset', preset: 'today' }, { ...ctx, nowMs: ctx.nowMs + 86400000 }).start).toBe(today.end)
  })
  it.each([
    ['yesterday', day(2026, 3, 7), day(2026, 3, 8)],
    ['tomorrow', day(2026, 3, 9), day(2026, 3, 10)],
    ['thisWeek', day(2026, 3, 8), day(2026, 3, 15)],
    ['thisMonth', day(2026, 3, 1), day(2026, 4, 1)],
    ['thisYear', day(2026, 1, 1), day(2027, 1, 1)],
  ] as const)('lowers %s from Sunday with calendar boundaries', (preset, first, next) => {
    expect(lower({ kind: 'datePreset', preset })).toEqual({ start: serial(first), end: serial(next), complement: false })
  })
  it.each([
    ['America/New_York', '2026-11-01T17:00:00Z', day(2026, 11, 1)],
    ['Australia/Sydney', '2026-10-04T01:00:00Z', day(2026, 10, 4)],
    ['Australia/Sydney', '2026-04-05T02:00:00Z', day(2026, 4, 5)],
    ['America/New_York', '2026-01-01T02:00:00Z', day(2025, 12, 31)],
  ] as const)('uses local days across DST/zone boundary in %s', (timeZone, now, first) => {
    const result = lowerDateCondition({ kind: 'datePreset', preset: 'today' }, { ...ctx, timeZone, nowMs: Date.parse(now) })
    expect(result.start).toBe(serial(first))
    expect(result.end - result.start).toBe(1)
  })
  it('handles static operators and inclusive end-day fractions', () => {
    const first = day(2026, 3, 8), second = day(2026, 3, 9), start = serial(first)
    for (const [op, at, expected] of [['eq', start + 0.5, true], ['neq', start + 0.5, false], ['neq', start - 1, true], ['before', start, false], ['before', start - 0.1, true], ['after', start + 0.9, false], ['after', start + 1, true], ['between', start + 1.99, true], ['between', start + 2, false]] as const) {
      expect(matchesDateInterval(at, lower({ kind: 'date', op, first, second })), `${op} ${at}`).toBe(expected)
    }
    expect(matchesDateInterval(NaN, lower({ kind: 'date', op: 'neq', first }))).toBe(false)
  })
  it('preserves fictitious serial 60, 1904, and maximum-year end boundary', () => {
    expect(lower({ kind: 'date', op: 'eq', first: day(1900, 2, 29) })).toEqual({ start: 60, end: 61, complement: false })
    expect(lowerDateCondition({ kind: 'date', op: 'eq', first: day(1904, 1, 1) }, { ...ctx, dateSystem: '1904' })).toEqual({ start: 0, end: 1, complement: false })
    expect(lower({ kind: 'date', op: 'eq', first: day(9999, 12, 31) }).end).toBe(serial(day(9999, 12, 31)) + 1)
  })
  it('resolves the 1900 day-zero clock token as real December 1899 and 1904 negative days', () => {
    const early = { ...ctx, timeZone: 'UTC', nowMs: Date.parse('1899-12-31T12:00:00Z') }
    expect(lowerDateCondition({ kind: 'datePreset', preset: 'today' }, early)).toEqual({ start: 0, end: 1, complement: false })
    expect(lowerDateCondition({ kind: 'datePreset', preset: 'thisMonth' }, early).start).toBe(serial(day(1899, 12, 1)))
    expect(lowerDateCondition({ kind: 'datePreset', preset: 'today' }, { ...early, dateSystem: '1904' }).start).toBe(serial(day(1899, 12, 31), '1904'))
  })
  it('rejects invalid/reversed calendars and invalid clock/semantics', () => {
    for (const first of [day(2026, 2, 29), day(2026, 13, 1), day(2026, 1, 0), day(10000, 1, 1), day(2026.5, 1, 1)]) expect(() => lower({ kind: 'date', op: 'eq', first })).toThrow(RangeError)
    expect(() => lower({ kind: 'date', op: 'between', first: day(2026, 2, 2), second: day(2026, 2, 1) })).toThrow(/reversed/i)
    expect(() => lower({ kind: 'date', op: 'between', first: day(2026, 2, 2) })).toThrow(RangeError)
    expect(() => lowerDateCondition({ kind: 'date', op: 'eq', first: day(1900, 2, 29) }, { ...ctx, dateSystem: '1904' })).toThrow(RangeError)
    for (const nowMs of [NaN, Infinity]) expect(() => lowerDateCondition({ kind: 'datePreset', preset: 'today' }, { ...ctx, nowMs })).toThrow(RangeError)
    expect(() => lowerDateCondition({ kind: 'datePreset', preset: 'today' }, { ...ctx, timeZone: 'invalid-zone' })).toThrow(RangeError)
    expect(() => lower({ kind: 'datePreset', preset: 'bad' as DatePreset })).toThrow(RangeError)
  })
})
