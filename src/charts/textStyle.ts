/** Font descriptions shared by measurement and paint, without a painter dependency. */
import type { LabelRun } from './labels'
export const chartFont = (px: number, family: string) => `${px}px ${family}`
export const labelFont = (run: LabelRun | undefined, px: number, family: string) =>
  `${run?.italic ? 'italic ' : ''}${run?.bold ? 'bold ' : ''}${chartFont((run?.sizePt ?? px * 72 / 96) * 96 / 72, family)}`
