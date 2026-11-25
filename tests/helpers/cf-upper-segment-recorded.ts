/**
 * The upper-segment registration (stage2-scale-upper-segment, 2ee3735a; scale:upper-segment and scale:upper-segment-near-top
 * at index-position) paints upper-segment cells that every earlier case's tests record as withheld. Those tests pin their
 * records as of their own captures: this installs, around each test of the calling file, an invalid proposal for both ids
 * (the test seams), so production and the justifications behave exactly as before the registration. The registered state is
 * pinned by tests/xlsx-cf-scale-upper-segment-case.test.ts and tests/xlsx-cf-scale-upper-segment-readings.test.ts.
 */
import { afterEach, beforeEach } from 'vitest'
import { CF_UPPER_NEAR_TOP_CAP, CF_UPPER_SEGMENT_CAP, injectCfUpperNearTopReadings, injectCfUpperSegmentReadings, type CfUpperSegmentReading } from '../../src/xlsx/cf/colorScale'

/** Install the pre-registration state (both ids unregistered) around every test of the calling file. */
export function recordedBeforeUpperSegmentRegistration(): void {
  beforeEach(() => unregisterUpperSegment())
  afterEach(() => { injectCfUpperSegmentReadings(null); injectCfUpperNearTopReadings(null) })
}

/** Both upper-segment ids unregistered through the seams (an invalid proposal shadows the production table). */
export function unregisterUpperSegment(): void {
  injectCfUpperSegmentReadings({ [CF_UPPER_SEGMENT_CAP]: 'unregistered' as CfUpperSegmentReading })
  injectCfUpperNearTopReadings({ [CF_UPPER_NEAR_TOP_CAP]: 'unregistered' as CfUpperSegmentReading })
}
