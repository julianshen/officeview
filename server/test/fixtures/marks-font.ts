/**
 * A pinned face for fixtures whose geometry depends on glyph shapes: stacked combining marks and small raised glyphs.
 *
 * `sans-serif` is the platform's font, and stacked marks are not portable through it (measured, @napi-rs/canvas,
 * 40 × U+030D after 'secret' at 20 px): macOS stacks them about 185 px tall; Linux (DejaVu Sans, or Liberation Sans
 * once the font pack is registered earlier in the process) does not stack them at all (ink 19–21 px). Liter
 * (`tests/fixtures/fonts`, OFL) positions U+0301 itself, mark on mark, with byte-identical metrics on both platforms,
 * with or without the pack: 40 × U+0301 after 'secret' are 201.64 px tall at 20 px, and its degree sign is 7 px tall
 * at 24 px. Every glyph a fixture draws in it must be Liter's (a missing one falls back to the platform font).
 */
import { GlobalFonts } from '@napi-rs/canvas'
import { resolve } from 'node:path'

/** The family the face is registered under (unique, so no platform or pack face can answer to it). */
export const MARKS_FAMILY = 'OV Test Marks'
/** A combining mark Liter stacks (U+0301 COMBINING ACUTE ACCENT). */
export const STACKING_MARK = '́'

if (!GlobalFonts.registerFromPath(resolve(import.meta.dir, '../../../tests/fixtures/fonts/Liter-Regular.ttf'), MARKS_FAMILY))
  throw new Error(`could not register ${MARKS_FAMILY}`)
