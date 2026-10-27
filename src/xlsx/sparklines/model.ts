export type SparklineType = 'line' | 'column' | 'stacked'
export interface SparklineArea { sheet?: string; r0: number; c0: number; r1: number; c1: number }
export interface XlsxSparkline { host: { row: number; col: number } | undefined; data: SparklineArea | undefined; raw: { f?: string; sqref?: string } }
export type SparklineDiagnostic = { kind: 'unsupported' | 'parse'; feature: string; message: string }
export interface XlsxSparklineGroup {
  type: SparklineType
  lineWeightPt: number
  displayEmptyCellsAs: 'gap' | 'zero' | 'span'
  markers: boolean; high: boolean; low: boolean; first: boolean; last: boolean; negative: boolean
  displayXAxis: boolean; displayHidden: boolean; dateAxis: boolean; rightToLeft: boolean
  minAxisType: 'individual' | 'group' | 'custom'; maxAxisType: 'individual' | 'group' | 'custom'
  manualMin?: number; manualMax?: number
  colors: Partial<Record<'series' | 'negative' | 'axis' | 'markers' | 'first' | 'last' | 'high' | 'low', string>>
  dateRange?: SparklineArea
  sparklines: XlsxSparkline[]
  refused?: string
}
