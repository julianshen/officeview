/**
 * FC-T2: the receiving grant of a drawings read (`replyByteLimit`). Main
 * reserves the grant on its ledger before it sends the command; the worker
 * never posts a reply to that command which measures more than the grant
 * (measurePayloadBytes, the exact measure main charges the received clone
 * with). A reply over the grant is replaced by a payload-free quota reply that
 * names its measured size, so main can ask once more with a grant that fits.
 * Every posted reply carries its own exact size (`replyBytes`, counted in it),
 * so main adopts that much of the grant without walking the reply.
 * Main-safe: protocol helpers and the pure measure only.
 */
import type { WorkerReply } from '../xlsx/view/contracts'
import { measurePayloadBytes } from '../xlsx/view/payloadBytes'
import { CANCEL_MESSAGES, reply, type CommandOf } from './viewerProtocol'

/** The fixed-shape reply that replaces a drawings reply over its grant. */
export function replyLimitRefusal(command: CommandOf<'read'>, requiredBytes: number | undefined, limitBytes: number): WorkerReply<'read'> {
  // No size when the reply was too large to measure: asking again cannot help.
  return reply(command, command.owner, { kind: 'quota', resource: 'memory', ...(requiredBytes === undefined ? {} : { requiredBytes }), limitBytes })
}
/** `out` with `replyBytes`, its exact size including that field, when
 * `admit` accepts the measuring workspace and the size is within `limit`. A
 * reply over the grant, or one the pool cannot measure now, is replaced by the
 * payload-free quota reply (sized too; no requiredBytes when unmeasurable). */
export function admitReply(command: CommandOf<'read'>, out: WorkerReply, limit: number, admit: (bytes: number) => boolean): WorkerReply {
  const sized = { ...out, replyBytes: 0 }
  const bytes = measurePayloadBytes(sized, admit)
  if (bytes !== undefined && bytes <= limit) { sized.replyBytes = bytes; return sized }
  return withReplyBytes(replyLimitRefusal(command, bytes, limit))
}
/** A small payload-free reply with its exact size stamped. */
function withReplyBytes(out: WorkerReply): WorkerReply {
  const sized = { ...out, replyBytes: 0 }
  sized.replyBytes = measurePayloadBytes(sized)!
  return sized
}
/** The smallest grant the worker can always honour for `command`: its largest
 * payload-free terminal (the quota refusal, or a cancellation with its reason). */
export function replyLimitEnvelope(command: CommandOf<'read'>): number {
  let bytes = withReplyBytes(replyLimitRefusal(command, 0, 0)).replyBytes!
  for (const [code, message] of Object.entries(CANCEL_MESSAGES)) {
    bytes = Math.max(bytes, withReplyBytes(reply(command, command.owner, { kind: 'cancelled' }, [{ code, message }])).replyBytes!)
  }
  return bytes
}
