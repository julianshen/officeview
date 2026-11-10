/**
 * `officeview/core` — the React-free entry (spec §8.1). Built in its own Vite
 * pass (`vite.core.config.ts`); its static-import closure must stay under the
 * 38,720 B eager ceiling and reach no parser or painter module
 * (`bun run gates:core`). Everything heavier loads through
 * `loadStreamRuntime()`.
 */
export { isStreamSource, StreamSourceMisuseError } from './stream/source'
export type { StreamSource } from './stream/source'

/** Loads the lazy stream runtime chunk. */
export function loadStreamRuntime(): Promise<typeof import('./stream/index')> {
  return import('./stream/index')
}
