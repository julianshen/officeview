/**
 * Browser stand-in for `node:crypto`, aliased ONLY by vite.harness.config.ts so
 * the dev harness can reuse the S0.2 fixture builders from
 * tests/helpers/xlsx-viewer-fixtures.ts in a page/worker. Those builders only
 * need `createHash` for their canonical-sample digests, which the harness never
 * calls; a call fails loudly instead of returning a wrong digest.
 */
export function createHash(algorithm: string): never {
  throw new Error(`node:crypto createHash(${algorithm}) is not available in the browser harness`)
}
