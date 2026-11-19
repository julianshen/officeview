/**
 * Piece 5 Task 11: records the wire bundles the remote-show media tests play
 * (`tests/stream/remote-show-media.test.tsx`, `remote-show-presenter-media.test.tsx`).
 *
 *   cd server && bun ../scripts/stream-p5/record-media-show-fixtures.ts
 *
 * For each piece-5 media deck (`src/testdata/show-decks.ts`) it runs a real harness session with media on (the
 * ffmpeg probe must pass: the recorder refuses to write a fixture from a media-off server), and writes, as
 * `scripts/record-show-fixtures.ts` does, every show-plan reply payload and every still unit payload to
 * `tests/fixtures/show-plan/media-<name>.bin`. Header (op 0) JSON: `{ units, visible, list }`, `list` the manifest's
 * `units.list` (its `anim` and `media` flags).
 *
 * Reproducible: Layer-1 is off, and each plan's session media ids (random per session, D2) are renamed to
 * deterministic ones (`M<deck>_<n>`, padded to 22 characters) before the plan message is re-encoded.
 */
import { mkdirSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import '../stream-s0/napi-preload'

const { Client, disposePool, start } = await import('../../server/test/harness')
const { decodeShowMessage, encodeShowMessage, MEDIA_ID, SHOW_PLAN_LIMITS } = await import('../../src/stream/show-plan')
const decks = await import('../../src/testdata/show-decks')
type ShowPlanRecord = import('../../src/stream/show-plan').ShowPlanRecord

const OUT = resolve(import.meta.dir, '../../tests/fixtures/show-plan')

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

/** `n`-th deterministic media id of deck `tag` (22 base64url characters). */
const fixedId = (tag: string, n: number) => `M${tag}_${n}`.padEnd(22, 'x')

async function record(name: string, doc: Uint8Array): Promise<Uint8Array> {
  const h = await start({ doc, policy: { contentPolicy: 'full', forensicMark: false }, media: {} })
  const c = new Client(h.url)
  try {
    const m = await c.hello(); await c.ready()
    if (m.can.media !== true) throw new Error(`${name}: the server has media off (ffmpeg probe failed?); refusing to record`)
    const list = m.units.list
    const visible = list.map((u, i) => (u.hidden ? -1 : i)).filter((i) => i >= 0)
    const recs: Rec[] = [{ op: 0, unit: 0, a: 0, b: 0, bytes: new TextEncoder().encode(JSON.stringify({ units: list.length, visible, list })) }]
    const ids = new Map<string, string>()
    let reqId = 100
    for (const u of visible) {
      let planned = false
      if (list[u]!.anim || list[u]!.media) {
        const r = await c.showPlan(u, 1, 'full')
        if (!r.end) {
          planned = true
          const [head, ...rest] = r.messages
          const plan = (decodeShowMessage(head!, { unit: u, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }) as { plan: ShowPlanRecord }).plan
          let json = JSON.stringify(plan)
          for (const rec of plan.media ?? []) if (!ids.has(rec.mediaId)) ids.set(rec.mediaId, fixedId(name.slice(0, 6), ids.size))
          for (const [from, to] of ids) json = json.split(from).join(to)
          const fixed = JSON.parse(json) as ShowPlanRecord
          for (const rec of fixed.media ?? []) if (!MEDIA_ID.test(rec.mediaId)) throw new Error(`${name}: bad fixed id ${rec.mediaId}`)
          recs.push({ op: 1, unit: u, a: 0, b: 0, bytes: encodeShowMessage({ tag: 1, plan: fixed }) })
          for (const p of rest) recs.push({ op: 1, unit: u, a: 0, b: 0, bytes: p })
          for (const [k, s] of fixed.steps.entries()) if (s.eligibility === 'composited') throw new Error(`${name}: unit ${u} step ${k} is composited (not recorded)`)
        } else console.log(`${name}: unit ${u} plan refused ${r.end.code}`)
      }
      if (!planned) {
        const id = reqId++
        c.unit(id, u, 1)
        const r = await c.unitMessages(id)
        const last = r.messages[r.messages.length - 1]!
        if (last.kind !== 3) throw new Error(`${name}: unit ${u}: expected a kind-3 unit message, got ${last.kind}`)
        recs.push({ op: 4, unit: u, a: 0, b: 0, bytes: last.payload })
      }
    }
    return encode(recs)
  } finally { c.close(); await h.stop() }
}

/** Two slides, each with four audio shapes begun with the slide (eight media parts across current + next). */
async function eightPartsDeck(): Promise<Uint8Array> {
  const node = (id: number, spid: number) => `<p:audio><p:cMediaNode vol="80000"><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cMediaNode></p:audio>`
  const files = ['a.m4a', 'a.mp3', 'a.wav', 'a.m4a'] as const
  const slide = () => ({ media: files.map((file, i) => ({ spid: 3 + i, file })), timing: files.map((_, i) => node(3 + i, 3 + i)).join('') })
  return decks.mediaDeck({ slides: [slide(), slide()] })
}

/** One slide: a narration over an undecodable part (`badhdr.m4a`, a Task 3 fixture): its record ships refused. */
const refusedNarrationDeck = () => decks.mediaDeck({ slides: [{ media: [{ spid: 3, file: 'badhdr.m4a' as 'a.m4a' }], timing: '<p:audio isNarration="1"><p:cMediaNode vol="80000"><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="3"/></p:tgtEl></p:cMediaNode></p:audio>' }] })

const WHICH: Array<[string, () => Promise<Uint8Array>]> = [
  ['narration', decks.narrationDeck],
  ['effect-sound', decks.effectSoundDeck],
  ['transition-sound', decks.transitionSoundDeck],
  ['cross-slide', decks.crossSlideSoundDeck],
  ['video', decks.videoDeck],
  ['unrelated', () => decks.mediaEndRevealDeck('unrelated')],
  ['eight-parts', eightPartsDeck],
  ['refused', refusedNarrationDeck],
]

mkdirSync(OUT, { recursive: true })
for (const [name, deck] of WHICH) {
  const bytes = await record(name, await deck())
  writeFileSync(resolve(OUT, `media-${name}.bin`), bytes)
  console.log(`media-${name}.bin: ${bytes.length} B`)
}
disposePool()
process.exit(0)
