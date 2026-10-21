import type { ChartPoint, ChartPoints } from '../../src/charts/model'
import { pointList } from '../../src/charts/points'
/** Point states for assertions: a number is a value, undefined blank, '#…' an error code, 'unavailable' unavailable. */
export const pts = (...items: Array<number | undefined | string>): ChartPoint[] => items.map(item =>
  typeof item === 'number' ? { state: 'value', v: item }
    : item === undefined ? { state: 'blank' }
      : item === 'unavailable' ? { state: 'unavailable' } : { state: 'error', code: item })
/** A role's compact storage as the `ChartPoint` union (the accessor view), for assertions. */
export const list = <T>(points: ChartPoints | T): ChartPoint[] | T => points && typeof points === 'object' && 'states' in (points as object) ? pointList(points as ChartPoints) : points as T
/** A deep copy with typed arrays as plain arrays: compares chart models across realms (a jsdom test realm
 * versus the structured clone a worker reply produces, whose typed arrays have the Node realm's constructors). */
export const plain = (value: unknown): unknown => ArrayBuffer.isView(value) ? Array.from(value as unknown as ArrayLike<number>)
  : Array.isArray(value) ? value.map(plain)
    : value && typeof value === 'object' ? Object.fromEntries(Object.entries(value).map(([k, v]) => [k, plain(v)])) : value
