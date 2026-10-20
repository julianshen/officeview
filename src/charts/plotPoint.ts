import type { ChartFamily, ChartGrouping, ChartPoint } from './model'
/** Spec 5.8: state identity survives until the family's plotting decision. */
export function plotPoint(point: ChartPoint, family: ChartFamily, grouping: ChartGrouping, blanks: 'gap'|'zero'|'span' = 'gap', naAsBlank=false): number|'gap'|'span' {
  if(point.state==='value')return point.v
  if(grouping==='stacked'||grouping==='percentStacked')return 0
  if(family==='bar'||family==='pie'||family==='doughnut')return 'gap'
  if(family==='area')return blanks==='gap'?'gap':0
  if(point.state==='unavailable')return 'gap'
  if(point.state==='error' && !(point.code==='#N/A'&&naAsBlank))return 'span'
  return blanks==='zero'?0:blanks
}
