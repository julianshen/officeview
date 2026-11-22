/**
 * Stage 4a (R9): the session clock `(day, timeZone)`, fixed for a sid at open. A tile reads no clock: TODAY() and
 * date conditional formats see `clockEpochMs` (12:00 local on `day`) and `timeZone`, never a client `nowMs`.
 */
export interface SheetClock { /** `YYYY-MM-DD`, the calendar date in `timeZone`. */ day: string; timeZone: string }

const formatter = (timeZone: string) => new Intl.DateTimeFormat('en-US', {
  timeZone, calendar: 'gregory', numberingSystem: 'latn', hourCycle: 'h23',
  year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit',
})

/** The wall-clock fields of `ms` in `timeZone`. */
function fieldsAt(f: Intl.DateTimeFormat, ms: number) {
  const parts: Record<string, number> = {}
  for (const p of f.formatToParts(ms)) if (p.type !== 'literal') parts[p.type] = Number(p.value)
  return { y: parts.year!, m: parts.month!, d: parts.day!, h: parts.hour! % 24, min: parts.minute!, s: parts.second! }
}
const pad = (n: number, w = 2) => String(n).padStart(w, '0')

/** The session clock at `nowMs`. Throws `RangeError` for a zone `Intl` rejects. The zone string is kept verbatim. */
export function sheetClockAt(nowMs: number, timeZone: string): SheetClock {
  if (typeof timeZone !== 'string') throw new RangeError('timeZone must be a string')
  const { y, m, d } = fieldsAt(formatter(timeZone), nowMs)
  return { day: `${pad(y, 4)}-${pad(m)}-${pad(d)}`, timeZone }
}

/** The instant of 12:00 on `c.day` in `c.timeZone`. */
export function clockEpochMs(c: SheetClock): number {
  const [y, m, d] = c.day.split('-').map(Number) as [number, number, number]
  const f = formatter(c.timeZone)
  const target = Date.UTC(y, m - 1, d, 12)
  let t = target
  // Two passes settle the zone offset (a DST change near noon moves it once).
  for (let i = 0; i < 2; i++) {
    const w = fieldsAt(f, t)
    t += target - Date.UTC(w.y, w.m - 1, w.d, w.h, w.min, w.s)
  }
  return t
}
