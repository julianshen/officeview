/** DOM-free slide preparation (piece 4 spec §4 item 1): shared by the local stage and the stream server. */
export {
  baseGeometry, computeTextTiming, hitRegionsOf, layoutTextBases, runRegionsOf, prepareSlide, rasterLayers, rasterPreparedLayer, targetsOf,
  type PrepareInput, type PreparedSlide,
} from './slide'
export { timingTargets } from '../engine/targets'
