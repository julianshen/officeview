/**
 * PR #78 round 13: what one show job (plan, track or Morph) is admitted for, and so the preparation grant its
 * build holds on the worker. Dependency-free, so the socket layer (`ws.ts`) and the build share one formula.
 */

/** The most one show job is admitted for. */
export const SHOW_BUILD_BYTES = 64 * 1024 * 1024

/**
 * Four page surfaces at the bucket (`w·h·4·scale²·4`), at most `SHOW_BUILD_BYTES`: the bytes `onShow` admits
 * for a job on a slide of `size` (slide px). The build leases every layer raster (a growth re-raster included)
 * and every masked copy of one against a ledger of this size before allocating it; a lease that does not fit
 * refuses the slide to its still (`show-budget`), and a Morph pair to its fallback.
 */
export const showBuildGrant = (size: { w: number; h: number }, scale: number) => Math.min(SHOW_BUILD_BYTES, size.w * size.h * 4 * scale ** 2 * 4)
