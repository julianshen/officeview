/**
 * FC-T2: the receiving grant of a drawings read (`replyByteLimit`). Main
 * reserves the grant on its ledger before it sends the command; the worker
 * never posts a reply to that command which measures more than the grant
 * (estimatePayloadBytes, the same measure main charges the received clone
 * with). A reply over the grant is replaced by a payload-free quota reply that
 * names its measured size, so main can ask once more with a grant that fits.
 * Main-safe: protocol helpers and the pure measure only.
 */
import type { WorkerReply } from '../xlsx/view/contracts'
import { estimatePayloadBytes } from '../xlsx/view/payloadBytes'
import { CANCEL_MESSAGES, reply, type CommandOf } from './viewerProtocol'

/** The fixed-shape reply that replaces a drawings reply over its grant. */
export function replyLimitRefusal(command: CommandOf<'read'>, requiredBytes: number, limitBytes: number): WorkerReply<'read'> {
  return reply(command, command.owner, { kind: 'quota', resource: 'memory', requiredBytes, limitBytes })
}
/** The smallest grant the worker can always honour for `command`: its largest
 * payload-free terminal (the quota refusal, or a cancellation with its reason). */
export function replyLimitEnvelope(command: CommandOf<'read'>): number {
  let bytes = estimatePayloadBytes(replyLimitRefusal(command, 0, 0))
  for (const [code, message] of Object.entries(CANCEL_MESSAGES)) {
    bytes = Math.max(bytes, estimatePayloadBytes(reply(command, command.owner, { kind: 'cancelled' }, [{ code, message }])))
  }
  return bytes
}
