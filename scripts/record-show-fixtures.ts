/**
 * Piece 4 Task 9 (T5 gate, spec §9 item 2): records the wire bundles the
 * no-package prototype (`tests/stream/t5-no-package.test.ts`) plays.
 *
 *   cd server && bun ../scripts/record-show-fixtures.ts
 *
 * For `flyInDeck` and `morphDeck` it runs a real harness session (the stream
 * server on port 0, the frame-decoding client) and writes every reply payload
 * the show would receive to `tests/fixtures/show-plan/{fly,morph}.bin`:
 *
 * - an `anim` slide: its `show-plan` kind-6 messages, then the kind-6
 *   messages of the `show-track` of every `composited` step;
 * - a slide entered by Morph: the `show-morph` kind-6 messages of the pair
 *   (previous slide, slide, previous slide's step count);
 * - a still slide (not `anim`, or a plan the server refused): its piece-3
 *   unit message (kind 3, a raster unit body).
 *
 * File format (little endian), a sequence of records:
 *   u8 op · u32 unit · u32 a · u32 b · u32 len · len bytes
 * op 0 — header JSON `{ units, visible }` (unit = a = b = 0);
 * op 1 — a show-plan message (`unit`); op 2 — a show-track message (`unit`, a = step);
 * op 3 — a show-morph message (`unit` = to, a = from, b = fromStep);
 * op 4 — a still: the raw kind-3 unit payload (`unit`).
 *
 * Layer-1 is off (`forensicMark: false`) so the committed bytes are
 * reproducible; the visible mark stays on.
 */
import { mkdirSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import './stream-s0/napi-preload'

const { Client, disposePool, start } = await import('../server/test/harness')
const { decodeShowMessage, SHOW_PLAN_LIMITS } = await import('../src/stream/show-plan')
const { flyInDeck, morphDeck } = await import('../src/testdata/show-decks')
type ShowPlanRecord = import('../src/stream/show-plan').ShowPlanRecord

const OUT = resolve(import.meta.dir, '../tests/fixtures/show-plan')

interface Rec { op: number; unit: number; a: number; b: number; bytes: Uint8Array }
function encode(recs: readonly Rec[]): Uint8Array {
  const size = recs.reduce((n, r) => n + 17 + r.bytes.length, 0)
  const out = new Uint8Array(size), v = new DataView(out.buffer)
  let at = 0
  for (const r of recs) {
    v.setUint8(at, r.op); v.setUint32(at + 1, r.unit, true); v.setUint32(at + 5, r.a, true); v.setUint32(at + 9, r.b, true); v.setUint32(at + 13, r.bytes.length, true)
    out.set(r.bytes, at + 17)
    at += 17 + r.bytes.length
  }
  return out
}

async function record(doc: Uint8Array): Promise<Uint8Array> {
  const h = await start({ doc, policy: { contentPolicy: 'full', forensicMark: false } })
  const c = new Client(h.url)
  try {
    const m = await c.hello(); await c.ready()
    const units = m.units.list
    const visible = units.map((u, i) => (u.hidden ? -1 : i)).filter((i) => i >= 0)
    const recs: Rec[] = [{ op: 0, unit: 0, a: 0, b: 0, bytes: new TextEncoder().encode(JSON.stringify({ units: units.length, visible })) }]
    const steps = new Map<number, number>()
    let reqId = 100
    const still = async (u: number) => {
      const id = reqId++
      c.unit(id, u, 1)
      const r = await c.unitMessages(id)
      const last = r.messages[r.messages.length - 1]!
      if (last.kind !== 3) throw new Error(`unit ${u}: expected a kind-3 unit message, got ${last.kind}`)
      recs.push({ op: 4, unit: u, a: 0, b: 0, bytes: last.payload })
    }
    for (const u of visible) {
      let plan: ShowPlanRecord | undefined
      if (units[u]!.anim) {
        const r = await c.showPlan(u, 1, 'full')
        if (!r.end) {
          plan = (decodeShowMessage(r.messages[0]!, { unit: u, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }) as { plan: ShowPlanRecord }).plan
          for (const p of r.messages) recs.push({ op: 1, unit: u, a: 0, b: 0, bytes: p })
          for (const [k, s] of plan.steps.entries()) {
            if (s.eligibility !== 'composited') continue
            const t = await c.showTrack(u, k, 1, 'full')
            if (t.end) throw new Error(`unit ${u} step ${k}: track refused ${t.end.code}`)
            for (const p of t.messages) recs.push({ op: 2, unit: u, a: k, b: 0, bytes: p })
          }
        }
      }
      if (!plan) await still(u)
      steps.set(u, plan?.steps.length ?? 0)
      const from = visible[visible.indexOf(u) - 1]
      if (plan?.show.transition?.element === 'morph' && from !== undefined) {
        const r = await c.showMorph(from, u, steps.get(from) ?? 0, 1, 'full')
        if (r.end) throw new Error(`morph ${from} → ${u} refused ${r.end.code}`)
        for (const p of r.messages) recs.push({ op: 3, unit: u, a: from, b: steps.get(from) ?? 0, bytes: p })
      }
    }
    return encode(recs)
  } finally { c.close(); h.stop() }
}

mkdirSync(OUT, { recursive: true })
for (const [name, deck] of [['fly', flyInDeck], ['morph', () => morphDeck()]] as const) {
  const bytes = await record(await deck())
  writeFileSync(resolve(OUT, `${name}.bin`), bytes)
  console.log(`${name}.bin: ${bytes.length} B`)
}
disposePool()
process.exit(0)
