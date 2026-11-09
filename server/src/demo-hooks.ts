/**
 * Demo `authorize` / `resolveDocument` (§6.1 hooks). Every viewer is `demo`
 * with the protected defaults; `?relaxed=1` asks for an audited commands
 * session ONLY when the server runs with `--demo`. Documents come from one
 * directory by bare file name (`?doc=deck.pptx`); anything else resolves to
 * nothing. A real deployment replaces both hooks with its own auth and store.
 */
import { readFile, realpath } from 'node:fs/promises'
import { basename, dirname, resolve } from 'node:path'
import type { Authorize, ResolveDocument } from './hooks'

const NAME = /^[\w][\w.-]*\.(docx|xlsx|pptx)$/

export function safeDocPath(docDir: string, name: string | null): string | undefined {
  if (!docDir || !name || name.length > 255 || !NAME.test(name) || name.includes('..') || basename(name) !== name) return undefined
  const root = resolve(docDir), at = resolve(root, name)
  return dirname(at) === root ? at : undefined
}

export function demoHooks(o: { docDir: string; demo: boolean }): { authorize: Authorize; resolveDocument: ResolveDocument } {
  return {
    authorize: async (req) => ({
      viewerId: 'demo',
      policy: o.demo && req.url.searchParams.get('relaxed') === '1' ? { projection: 'relaxed' } : {},
    }),
    resolveDocument: async (req) => {
      const name = req.url.searchParams.get('doc')
      const at = safeDocPath(o.docDir, name)
      if (!at) return null
      try {
        // Containment on REAL paths: a symlink inside the directory must not
        // reach a file outside it (review r2-10).
        const [root, real] = await Promise.all([realpath(o.docDir), realpath(at)])
        if (dirname(real) !== root) return null
        return { bytes: new Uint8Array(await readFile(real)), name: name! }
      } catch { return null }
    },
  }
}
