/**
 * T12 render worker: loads a small fixed set of corpus units once, reports its
 * canvas backend, then paints `{ id, key }` requests and replies with the time.
 * Started by t12-admission.ts with the napi preload.
 */
declare const self: Worker
import { listCorpus, loadUnits, renderUnit } from './corpus'
import { resolveBackend } from './results'

const FILES = 6
const byKey = new Map<string, Awaited<ReturnType<typeof loadUnits>>[number]>()
const units: Array<{ key: string; bytes: number }> = []
const pick = listCorpus().filter((f) => !f.name.startsWith('poi-crash')).slice(0, FILES * 3)
for (const file of pick) {
  if (units.length >= FILES) break
  try {
    const loaded = await loadUnits(file)
    const unit = loaded[0]
    if (!unit) continue
    const key = `${file.name}#${unit.index}`
    byKey.set(key, unit)
    units.push({ key, bytes: Math.ceil(unit.widthPx) * Math.ceil(unit.heightPx) * 4 })
  } catch {
    // corrupt or unrenderable files are skipped for the synthetic load
  }
}
self.postMessage({ backend: resolveBackend(), units })

self.onmessage = async (e: MessageEvent<{ id: number; key: string }>) => {
  const t0 = performance.now()
  try {
    await renderUnit(byKey.get(e.data.key)!, 1)
    self.postMessage({ id: e.data.id, ms: performance.now() - t0, backend: resolveBackend() })
  } catch (error) {
    self.postMessage({ id: e.data.id, ms: performance.now() - t0, backend: resolveBackend(), error: (error as Error).message })
  }
}
