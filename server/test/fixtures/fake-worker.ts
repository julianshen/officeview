/**
 * A controllable stand-in for a Bun Worker, for the start gate and the pool's startup paths (`startWorker`'s `create`
 * seam, `PoolOptions.workerFactoryForTests`). Each instance records its real state: constructed, acknowledged (it sent
 * its first message), terminated (and whether that happened while it was still unacknowledged), closed, and the
 * messages it was sent.
 *
 * Scripts, in construction order (`next` defaults to `{ init: 0 }`):
 * - `init: n` — answers the first message (the pool's `init`) after n ms; `init: 'manual'` waits for `ack()`.
 * - `die: true` — closes on its own before it answers anything (a load that fails).
 * - `throws: true` — the constructor throws.
 * Once acknowledged, a job (any message with a numeric `id`) is answered at once with an ok reply.
 */
export type Script = { init?: number | 'manual'; die?: boolean; throws?: boolean }

export class FakeWorker {
  static all: FakeWorker[] = []
  static scripts: Script[] = []
  static reset() { FakeWorker.all = []; FakeWorker.scripts = [] }
  /** Settles every fake still alive (acknowledges and closes it), so no fake holds the process-wide gate. */
  static async release() {
    for (const w of FakeWorker.all) { if (!w.acked && !w.closed) w.ack(); if (!w.closed) w.close() }
    await Bun.sleep(0)
  }

  onmessage: ((e: MessageEvent) => void) | null = null
  acked = false
  closed = false
  terminateCalls = 0
  terminatedUnacked = false
  received: unknown[] = []
  jobs: unknown[] = []
  private closeListeners: Array<() => void> = []
  private readonly script: Script

  constructor(readonly url: string) {
    const script = FakeWorker.scripts.shift() ?? { init: 0 }
    if (script.throws) throw new Error('fake: worker construction failed')
    this.script = script
    FakeWorker.all.push(this)
    if (script.die) setTimeout(() => this.close(), 5)
  }

  postMessage(m: unknown) {
    if (this.closed) throw new DOMException('Worker has been terminated', 'InvalidStateError')
    this.received.push(m)
    const id = (m as { id?: unknown }).id
    if (!this.acked) {
      if (this.received.length === 1 && !this.script.die && this.script.init !== 'manual') setTimeout(() => this.ack(), this.script.init ?? 0)
      return
    }
    if (typeof id === 'number') this.answer(m)
  }

  /** Sends the first answer (the init reply), then answers every job it was sent meanwhile. */
  ack() {
    if (this.acked || this.closed) return
    this.acked = true
    const [first, ...rest] = this.received
    this.reply((first as { id?: number } | undefined)?.id ?? -2)
    for (const m of rest) if (typeof (m as { id?: unknown }).id === 'number') this.answer(m)
  }

  private answer(m: unknown) {
    this.jobs.push(m)
    this.reply((m as { id: number }).id)
  }

  private reply(id: number) {
    const data = { id, ok: true, result: null, stats: { parses: 0, pages: 0, docs: 0, probes: 0 } }
    queueMicrotask(() => { if (!this.closed) this.onmessage?.({ data } as MessageEvent) })
  }

  terminate() {
    this.terminateCalls++
    if (!this.acked) this.terminatedUnacked = true
    setTimeout(() => this.close(), 1)
  }

  close() {
    if (this.closed) return
    this.closed = true
    for (const f of this.closeListeners) f()
  }

  addEventListener(type: string, f: () => void) { if (type === 'close') this.closeListeners.push(f) }
}

export const fakeFactory = (url: string) => new FakeWorker(url) as unknown as Worker
