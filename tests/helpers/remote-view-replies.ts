import { ViewerClient, type ViewerPort } from '../../src/worker/viewerClient'
import { RemoteWorkbook } from '../../src/worker/viewerRemoteSource'
import { initialIdentity, initialWorkbookOwner } from '../../src/worker/viewerProtocol'
import { measurePayloadBytes } from '../../src/xlsx/view/payloadBytes'
import type { RowPermutationBand, SortSpec, ViewCommitDelta, WorkerCommand } from '../../src/xlsx/view/contracts'

/** Controlled message events: no import, rendering, polling or timer readiness. */
export function remoteViewReplies() {
  const commands: WorkerCommand[] = []
  const waiters: Array<{ kind: string; resolve(command: WorkerCommand): void }> = []
  const listeners = new Set<(event: { data: unknown }) => void>()
  const port: ViewerPort = {
    postMessage(message) {
      const command = message as WorkerCommand
      const index = waiters.findIndex(waiter => waiter.kind === command.kind)
      if (index >= 0) waiters.splice(index, 1)[0].resolve(command)
      else commands.push(command)
    },
    addEventListener(_type, listener) { listeners.add(listener) },
    removeEventListener(_type, listener) { listeners.delete(listener) },
  }
  const client = new ViewerClient(port)
  const remote = new RemoteWorkbook(client, { docId: 'D', sheets: [{ sheetId: 's', name: 'S' }], valuePolicy: 'cached', diagnostics: [], layouts: new Map() })
  const next = <K extends WorkerCommand['kind']>(kind: K) => {
    const index = commands.findIndex(command => command.kind === kind)
    return (index >= 0 ? Promise.resolve(commands.splice(index, 1)[0]) : new Promise<WorkerCommand>(resolve => waiters.push({ kind, resolve }))) as Promise<Extract<WorkerCommand, { kind: K }>>
  }
  const reply = (command: WorkerCommand, result: unknown) => {
    const data = { commandKind: command.kind, requestId: command.requestId, owner: 'owner' in command ? command.owner : undefined, diagnostics: [], kind: 'complete', result, replyBytes: 0 }
    // As the kernel does (spec §10.2): the worker stamps the reply's exact size.
    data.replyBytes = measurePayloadBytes(data)!
    for (const listener of listeners) listener({ data })
    return data
  }
  const delta = (viewGen: number, bands: readonly RowPermutationBand[]): ViewCommitDelta => ({
    owner: { ...initialIdentity('D', 's'), viewGen }, workbookOwner: { ...initialWorkbookOwner('D', [{ sheetId: 's' }]), globalViewGen: viewGen },
    bands, visibility: [], metrics: { rows: [], cols: [] }, counts: { visible: 1048576, total: 1048576 }, selection: { active: null, anchor: null, ranges: [] },
  })
  const sort = async (viewGen: number, bands: readonly RowPermutationBand[], spec: SortSpec | null = null) => {
    const pending = remote.sort(remote.identity('s'), spec, new AbortController().signal)
    if (spec) { const prepare = await next('prepareColumns'); reply(prepare, { kind: 'prepareColumns', cols: new Uint16Array([0]) }) }
    const command = await next('sort')
    const out = reply(command, { kind: 'view', delta: delta(viewGen, bands) })
    return { result: await pending, out }
  }
  return { remote, next, reply, delta, sort }
}
