/**
 * Stage 1d Task 6 / §16 T2 "local loads byte-for-byte unchanged when no
 * stream source is supplied".
 *
 * Paints every unit of every `corpus/` document through the LOCAL path —
 * `loadOfficeFile(bytes)` → `getPaintables` — with @napi-rs/canvas at scale 1
 * and hashes the RGBA. The source tree is a parameter, so the same script
 * measures this branch and the pre-streaming baseline (a worktree at
 * ca952c50, the merge-base with origin/main):
 *
 *   bun --preload ./scripts/stream-s0/napi-preload.ts scripts/stream-s1/local-identity.ts --src <tree> --write <out.json>
 *   bun --preload ./scripts/stream-s0/napi-preload.ts scripts/stream-s1/local-identity.ts --src . <baseline.json>
 *
 * Compare mode exits 1 on any differing unit. Embedded faces get the same
 * explicit throwing registrar on both sides (no system fallback difference).
 */
import { createHash } from 'node:crypto'
import { readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { createCanvas } from '@napi-rs/canvas'

const arg = (name: string) => { const i = process.argv.indexOf(name); return i === -1 ? undefined : process.argv[i + 1] }
const src = resolve(arg('--src') ?? '.')
const write = arg('--write')
const compareWith = write ? undefined : process.argv.slice(2).filter((a, i, all) => !a.startsWith('--') && all[i - 1] !== '--src').pop()
const MAX_RGBA = 64 * 1024 * 1024

const { loadOfficeFile } = await import(join(src, 'src/components/OfficeFile.tsx'))
const { getPaintables } = await import(join(src, 'src/render/paint.ts'))
const corpusDir = resolve(import.meta.dir, '../../corpus')
const files = readdirSync(corpusDir).filter((f) => /\.(docx|xlsx|pptx)$/.test(f)).sort()
const registerFont = (r: { face?: { family?: string } }) => { throw new Error(`no embedded font registrar (${r.face?.family ?? '?'})`) }

const out: Record<string, string[]> = {}
for (const name of files) {
  try {
    const doc = await loadOfficeFile(new Uint8Array(readFileSync(join(corpusDir, name))))
    const units = await getPaintables(doc, { registerFont })
    out[name] = units.map((p: { spec: { widthPx: number; heightPx: number }; paint(c: unknown): void }) => {
      const w = Math.max(1, Math.ceil(p.spec.widthPx - 1e-9)), h = Math.max(1, Math.ceil(p.spec.heightPx - 1e-9))
      if (w * h * 4 > MAX_RGBA) return `oversize:${w}x${h}`
      const ctx = createCanvas(w, h).getContext('2d')
      p.paint(ctx)
      return `${w}x${h}:${createHash('sha256').update(ctx.getImageData(0, 0, w, h).data).digest('hex')}`
    })
    ;(units as { dispose?: () => void }).dispose?.()
  } catch (e) { out[name] = [`error:${(e as Error).name}:${(e as Error).message.slice(0, 80)}`] }
}

if (write) {
  writeFileSync(write, `${JSON.stringify({ source: src, files: out }, null, 1)}\n`)
  console.log(`wrote ${Object.keys(out).length} files, ${Object.values(out).flat().length} units → ${write}`)
} else if (compareWith) {
  const base = (JSON.parse(readFileSync(compareWith, 'utf8')) as { files: Record<string, string[]> }).files
  let units = 0, differing = 0
  const diffs: string[] = []
  for (const name of new Set([...Object.keys(base), ...Object.keys(out)])) {
    const a = base[name] ?? [], b = out[name] ?? []
    for (let i = 0; i < Math.max(a.length, b.length); i++) { units++; if (a[i] !== b[i]) { differing++; diffs.push(`${name}#${i}: ${a[i]} → ${b[i]}`) } }
  }
  console.log(`files ${Object.keys(out).length}, units ${units}, differing ${differing}`)
  for (const d of diffs.slice(0, 20)) console.log(`  ${d}`)
  process.exit(differing ? 1 : 0)
} else {
  console.error('pass --write <out.json> or a baseline json to compare')
  process.exit(2)
}
