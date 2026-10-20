import type { ChartPoint } from '../../src/charts/model'
/** Point states for assertions: a number is a value, undefined blank, '#…' an error code, 'unavailable' unavailable. */
export const pts = (...items: Array<number | undefined | string>): ChartPoint[] => items.map(item =>
  typeof item === 'number' ? { state: 'value', v: item }
    : item === undefined ? { state: 'blank' }
      : item === 'unavailable' ? { state: 'unavailable' } : { state: 'error', code: item })
/** The plotted numbers of a point array (non-values read as undefined). */
export const nums = (points: readonly ChartPoint[] | undefined): Array<number | undefined> | undefined => points?.map(p => p.state === 'value' ? p.v : undefined)
