/**
 * Wait budget for the FIRST painted gridcell of an interactive viewer mount.
 *
 * Testing Library's default `waitFor` timeout (1000 ms) times a cold start,
 * not behaviour: a viewer's first gridcell waits on the lazy
 * InteractiveSheetDoc chunk (src/components/OfficeDoc.tsx `lazy(...)`), the
 * paged-workbook / worker-source dynamic imports, the dynamic `canvas` import
 * of the first paint, the in-process worker handshake and jsdom layout. Each
 * test file is a fresh module graph, so every file pays this once, inside its
 * first `waitFor`.
 *
 * Measured in a Linux container (DejaVu fonts, bun 1.3.10, 2 vitest workers):
 * the first gridcell lands in 0.9-1.4 s at 1 vCPU and 2.2-3.1 s at 0.5 vCPU,
 * with the accessible name already carrying the value ("A1, value alpha").
 * GitHub's ubuntu-latest runner ran the suite about 2x slower than an
 * unthrottled container and timed out at exactly these waits (the first
 * mount of each file); at 0.5 vCPU a later mount's first paint in the same
 * files also passed 1 s, so those waits use it too. 10 s keeps a
 * large margin over that while still failing fast on a viewer that never
 * paints. A test on vitest's default 5 s test timeout also takes
 * FIRST_PAINT_TEST, so the wait (and the rest of the test) fits inside it;
 * tests that already declare a longer timeout keep theirs.
 *
 * Use only for a mount's first paint (or a first cold gesture with the same
 * evidence), never to widen ordinary waits.
 */
export const FIRST_PAINT_WAIT = { timeout: 10_000 } as const
export const FIRST_PAINT_TEST = { timeout: 30_000 } as const
