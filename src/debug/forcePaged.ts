/**
 * S13 debug-only admission override for the e2e page (`?paged=1`). A small
 * file (such as the native Excel checkpoint workbook) is far below every
 * object cap, so it would never reach the worker-backed paged path. Lowering
 * only the object populated-cell cap to 1 makes admission choose paged for
 * any workbook with two or more cells, on the main thread (loadOfficeSource)
 * and in the viewer worker (ViewerKernel importLimits) alike, without
 * touching shipped code: the library has no such switch, and src/debug is
 * excluded from the build (tsconfig.build.json).
 */
import type { ImportLimitOverrides } from '../xlsx/storage/importSession'

export const FORCE_PAGED_LIMITS: ImportLimitOverrides = Object.freeze({ object: Object.freeze({ cells: 1 }) })

/** Worker `name` segments are `;`-separated (`caps:a,b`, `paged`). */
export const workerNameSegments = (name: string | undefined): string[] => (name ?? '').split(';').filter(Boolean)
