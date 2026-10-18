/**
 * S10.2 (codex round 5, coordinator decision): compiled-code bytes of the
 * page's main V8 heap WITHOUT a heap snapshot. A heap snapshot assigns DOM
 * node ids and serializes numbers, growing Blink's IdToNodeMap and V8's
 * smi_string_cache, which the following readings then see. A memory-infra
 * trace dump reads V8's per-space statistics instead.
 *
 * Code bytes = allocated objects in the main isolate's code and trusted spaces
 * (machine code, and bytecode/feedback/metadata in the trusted spaces):
 * code_space + code_large_object_space + trusted_space +
 * trusted_large_object_space. The dump covers every renderer process; the
 * page's is the one whose main-heap allocated bytes best match the page's own
 * Runtime.getHeapUsage usedSize taken just before (within MATCH_TOLERANCE),
 * otherwise the reading is null and the gate stays pending.
 */
import type { CDPSession } from 'playwright'

const CODE_SPACES = ['code_space', 'code_large_object_space', 'trusted_space', 'trusted_large_object_space']
const MATCH_TOLERANCE = 0.1

export interface CodeSpaceReading { codeBytes: number; mainHeapBytes: number; usedSize: number; pid: number }

export async function measureCodeSpace(cdp: CDPSession, usedSize: number): Promise<CodeSpaceReading | null> {
  const chunks: unknown[] = []
  const collect = (e: { value: unknown[] }) => { chunks.push(...e.value) }
  cdp.on('Tracing.dataCollected', collect)
  const complete = new Promise<void>(resolve => cdp.once('Tracing.tracingComplete', () => resolve()))
  try {
    await cdp.send('Tracing.start', { traceConfig: { includedCategories: ['disabled-by-default-memory-infra'], excludedCategories: ['*'], memoryDumpConfig: { triggers: [] } } as never, transferMode: 'ReportEvents' })
    const dump = await cdp.send('Tracing.requestMemoryDump' as never, { deterministic: false, levelOfDetail: 'detailed' } as never) as { success?: boolean }
    await cdp.send('Tracing.end')
    await complete
    if (!dump?.success) return null
  } finally { cdp.off('Tracing.dataCollected', collect) }
  const perPid = new Map<number, { code: number; heap: number }>()
  for (const raw of chunks) {
    const ev = raw as { pid?: number; args?: { dumps?: { allocators?: Record<string, { attrs?: Record<string, { value?: string }> }> } } }
    const allocators = ev.args?.dumps?.allocators
    if (!allocators || typeof ev.pid !== 'number') continue
    const entry = perPid.get(ev.pid) ?? { code: 0, heap: 0 }
    for (const [name, a] of Object.entries(allocators)) {
      const m = /^v8\/main\/heap\/([a-z_]+)$/.exec(name)
      if (!m) continue
      const size = parseInt(a.attrs?.allocated_objects_size?.value ?? '0', 16)
      if (!Number.isFinite(size)) continue
      entry.heap += size
      if (CODE_SPACES.includes(m[1])) entry.code += size
    }
    perPid.set(ev.pid, entry)
  }
  let best: CodeSpaceReading | null = null
  for (const [pid, e] of perPid) {
    if (!(e.heap > 0)) continue
    if (!best || Math.abs(e.heap - usedSize) < Math.abs(best.mainHeapBytes - usedSize)) best = { codeBytes: e.code, mainHeapBytes: e.heap, usedSize, pid }
  }
  return best && Math.abs(best.mainHeapBytes - usedSize) <= usedSize * MATCH_TOLERANCE ? best : null
}
