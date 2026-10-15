import { expect, it } from 'vitest'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

it('imports the repository fixture cache path in the default jsdom environment', async () => {
  const root = join(dirname(fileURLToPath(import.meta.url)), '..')
  await expect(import('../scripts/harness/viewerE2EFixtureCache')).resolves.toMatchObject({
    VIEWER_E2E_FIXTURE_DIR: join(root, '.tmp', 'viewer-e2e-fixtures-v1'),
  })
})
