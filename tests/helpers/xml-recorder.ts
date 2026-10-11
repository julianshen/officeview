/** Event recording and feeding strategies for the S1.1 tokenizer tests. */
import { PackageError } from '../../src/core/package-error'
import { XmlTokenizer, XstringDecoder, localName, type XmlEvents, type XmlLimits } from '../../src/core/xml-stream'

export type RecordedEvent =
  | { kind: 'open'; name: string; attrs: Array<[string, string]>; selfClosing: boolean }
  | { kind: 'close'; name: string }
  | { kind: 'text'; text: string }

/** Text segmentation is legitimately chunk-dependent, so adjacent text events are coalesced. */
export function recorder(): { events: RecordedEvent[]; handler: XmlEvents; maxSegment: number } {
  const events: RecordedEvent[] = []
  const state = { events, maxSegment: 0, handler: undefined as unknown as XmlEvents }
  state.handler = {
    open(name, names, values, count, selfClosing) {
      const attrs: Array<[string, string]> = []
      for (let i = 0; i < count; i++) attrs.push([names[i], values[i]])
      events.push({ kind: 'open', name, attrs, selfClosing })
    },
    close(name) { events.push({ kind: 'close', name }) },
    text(text) {
      state.maxSegment = Math.max(state.maxSegment, text.length)
      if (text.length === 0) return
      const last = events[events.length - 1]
      if (last?.kind === 'text') last.text += text
      else events.push({ kind: 'text', text })
    },
  }
  return state
}

export type Outcome = { ok: true; events: RecordedEvent[] } | { ok: false; code: string; message: string }

/** Feed `doc` cut at the given ascending positions. */
export function outcomeOf(doc: string, cuts: readonly number[], limits?: Partial<XmlLimits>): Outcome {
  const rec = recorder()
  try {
    const tok = new XmlTokenizer(rec.handler, limits)
    let at = 0
    for (const cut of cuts) { tok.write(doc.slice(at, cut)); at = cut }
    tok.write(doc.slice(at))
    tok.end()
    return { ok: true, events: rec.events }
  } catch (error) {
    if (error instanceof PackageError) return { ok: false, code: error.code, message: error.message }
    throw error
  }
}

export function everyN(doc: string, n: number): number[] {
  const cuts: number[] = []
  for (let i = n; i < doc.length; i += n) cuts.push(i)
  return cuts
}

export function randomCuts(doc: string, rand: () => number, maxPiece: number): number[] {
  const cuts: number[] = []
  for (let at = 0; ;) {
    at += 1 + Math.floor(rand() * maxPiece)
    if (at >= doc.length) return cuts
    cuts.push(at)
  }
}

/** Comparable form: errors compare by code only (design note §3.3). */
export function key(o: Outcome): string {
  return o.ok ? JSON.stringify(o.events) : `error:${o.code}`
}

/**
 * sharedStrings `si` consumer with the current parser's semantics
 * (parse.ts sharedStrings: t and r/t text, not rPh), stateful ST_Xstring
 * decoding per text segment. Test-only: the production consumer is S1.1b.
 */
export function sharedStringsConsumer(): { handler: XmlEvents; strings: string[] } {
  const strings: string[] = []
  const stack: string[] = []
  let current = ''
  let inSi = false
  let decoder: XstringDecoder | null = null
  const handler: XmlEvents = {
    open(qname, _n, _v, _c, selfClosing) {
      const name = localName(qname)
      if (name === 'si' && stack.length === 1) { inSi = true; current = '' }
      if (selfClosing) {
        if (name === 'si' && inSi) { strings.push(current); inSi = false }
        return
      }
      stack.push(name)
      if (name === 't' && inSi && isSiText(stack)) decoder = new XstringDecoder()
    },
    close() {
      const name = stack.pop()
      if (name === 't' && decoder !== null) { current += decoder.push('', true); decoder = null }
      if (name === 'si' && inSi) { strings.push(current); inSi = false }
    },
    text(text) {
      if (decoder !== null) current += decoder.push(text)
    },
  }
  return { handler, strings }
}

function isSiText(stack: string[]): boolean {
  // sst > si > t  or  sst > si > r > t (rPh/t is ignored like parse.ts)
  return (stack.length === 3 && stack[1] === 'si') || (stack.length === 4 && stack[1] === 'si' && stack[2] === 'r')
}
