/**
 * Interfaces of the chart display-format adapter and text measurement
 * (spec 5.2, 6.5). Interfaces only: A2 implements `ChartFormatter` over
 * `src/xlsx/cellFormat.ts`; stage-2 layout receives a `TextMeasure`.
 */

export interface ChartFormatter {
  /** Formats a value with a number format code (undefined = General) in the chart epoch. */
  format(v: number, code: string | undefined, epoch: '1900' | '1904'): string
}

export interface TextMeasure {
  width(text: string, font: string): number
  lineHeight(font: string): number
}
