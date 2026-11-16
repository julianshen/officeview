/**
 * What one show job (plan, track or Morph pair) is admitted for, and the preparation grant its build holds on the
 * worker. Dependency-free, so the socket layer (`ws.ts`), the pool and the build share one formula.
 *
 * Owner ruling (PR #78 round 13b, "reserve per worker"):
 * - at request time a show job is admitted for `showRequestBytes(grant)` against the session aggregate (96 MiB):
 *   its place in the queue and its encoded reply, never its preparation workspace;
 * - when the job holds its worker's slot (one job per worker at a time), the pool reserves the full grant against
 *   that worker's preparation reservation, released when the job ends. Server preparation memory is therefore at
 *   most workers × `SHOW_BUILD_BYTES`, and every canvas the build makes is charged to the grant as it is made.
 */

/** Page surfaces a job's preparation may hold at once (measured peaks: corpus plans ≤ 4.0, synthetic plans ≤ 6.3, layered Morph pairs ≤ 5.6). */
export const SHOW_GRANT_PAGES = 8

/**
 * The most one show job is granted, and so each worker's preparation bound: 128 MiB, the per-show decoded budget
 * (`SHOW_PLAN_LIMITS.showDecoded`). Eight page surfaces at a 2560 × 1440 bucket (a 16:9 slide at scale 2, 14.7 MB
 * each) are 118 MB, under it; a 3840 × 2160 bucket (the client's backing maximum) gets 4 pages.
 */
export const SHOW_BUILD_BYTES = 128 * 1024 * 1024

/** The request-time admission of a show job: one encoded kind-6 message (8 MiB), or the grant when smaller. */
export const SHOW_REQUEST_BYTES = 8 * 1024 * 1024
export const showRequestBytes = (grant: number) => Math.min(SHOW_REQUEST_BYTES, grant)

const page = (size: { w: number; h: number }, scale: number) => size.w * size.h * 4 * scale ** 2

/** A plan or track job on a slide of `size` (slide px) at `scale`: `SHOW_GRANT_PAGES` pages, at most `SHOW_BUILD_BYTES`. */
export const showBuildGrant = (size: { w: number; h: number }, scale: number) => Math.min(SHOW_BUILD_BYTES, SHOW_GRANT_PAGES * page(size, scale))

/** A Morph pair: `SHOW_GRANT_PAGES` pages of its larger slide (both slides' layers share them), at most `SHOW_BUILD_BYTES`. */
export const showMorphGrant = (from: { w: number; h: number }, to: { w: number; h: number }, scale: number) =>
  Math.min(SHOW_BUILD_BYTES, SHOW_GRANT_PAGES * Math.max(page(from, scale), page(to, scale)))
