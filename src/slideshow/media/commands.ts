/**
 * Piece 5 (plan D15, X10, X11): the media command whitelist, shared by the local player (`player.ts` `command()`)
 * and the server's plan build (`inertCommands`), so both read a `p:cmd` the same way.
 *
 * - `evt` + `onstopaudio` (case-insensitive) is the "stop previous sound" end condition.
 * - `call` + `playFrom(<seconds>)`, `play`, `resume`, `pause`, `stop`, `togglePause` drive one media node. The
 *   `call` verbs parse only for `cmdType === 'call'`, as `command()` requires.
 * - Anything else is `undefined`.
 *
 * The parse does not bound `playFrom`: local playback seeks wherever the file says. The server re-ships only
 * `0 ≤ seconds ≤ 1800` (the audio cap), in canonical form (`canonicalMediaCommand`). DOM-free and package-free.
 */
export type MediaCommand =
  | { verb: 'playFrom'; seconds: number }
  | { verb: 'play' | 'resume' | 'pause' | 'stop' | 'togglePause' }
  | { verb: 'onstopaudio' }

const PLAY_FROM = /^playFrom\(\s*([0-9]+(?:\.[0-9]+)?)\s*\)$/
const VERBS = ['play', 'resume', 'pause', 'stop', 'togglePause'] as const

export function parseMediaCommand(cmdType: string, text: string): MediaCommand | undefined {
  const t = text.trim()
  if (cmdType === 'evt') return t.toLowerCase() === 'onstopaudio' ? { verb: 'onstopaudio' } : undefined
  if (cmdType !== 'call') return undefined
  const from = PLAY_FROM.exec(t)
  if (from) return { verb: 'playFrom', seconds: Number(from[1]) }
  return (VERBS as readonly string[]).includes(t) ? { verb: t as (typeof VERBS)[number] } : undefined
}

/** `n` (finite, ≥ 0) as its shortest round-trip decimal, never in exponent form (`1e-7` → `0.0000001`). */
function plainDecimal(n: number): string {
  const s = String(n)
  const m = /^(\d)(?:\.(\d+))?e([+-]\d+)$/.exec(s)
  if (!m) return s
  const digits = m[1]! + (m[2] ?? ''), exp = Number(m[3])
  // d.ddd × 10^exp: the decimal point sits after `1 + exp` digits.
  const at = 1 + exp
  if (at <= 0) return `0.${'0'.repeat(-at)}${digits}`
  if (at >= digits.length) return digits + '0'.repeat(at - digits.length)
  return `${digits.slice(0, at)}.${digits.slice(at)}`
}

/** The text the server ships for `c`: it re-parses to `c` (`playFrom` seconds finite and ≥ 0). */
export function canonicalMediaCommand(c: MediaCommand): string {
  if (c.verb !== 'playFrom') return c.verb
  if (!Number.isFinite(c.seconds) || c.seconds < 0) throw new RangeError(`playFrom(${c.seconds}) has no canonical form`)
  return `playFrom(${plainDecimal(c.seconds)})`
}
