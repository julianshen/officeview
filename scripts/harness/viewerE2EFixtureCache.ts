import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

// Both browser drivers reuse the repository's ignored cache, independently of
// the caller's working directory.
export const VIEWER_E2E_FIXTURE_DIR = resolve(dirname(fileURLToPath(import.meta.url)), '..', '..', '.tmp', 'viewer-e2e-fixtures-v1')
