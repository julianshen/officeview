/** Lazy sparkline chunk. Importing it registers painting; eager code reaches it only through sparklineLoader. */
import { registerSparklinePainter } from '../sparklineLoader'
import { paintSparklinesIn } from './paint'

registerSparklinePainter(paintSparklinesIn)
export { paintSparklinesIn }
