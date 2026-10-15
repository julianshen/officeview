import type { Page } from '@playwright/test'

export interface SortOutcome {
  outcome: 'completed' | 'refused' | 'timeout'
  elapsedMs: number
  status: string
  ariaSort: string | null
}

type SortWatch = { result: Promise<Omit<SortOutcome, 'elapsedMs'>>; dispose(): void }
type SortWindow = Window & { __viewerPerfSortWatch?: SortWatch }

/** Install before input so an immediate refusal cannot be missed while awaiting its acknowledgment. */
export async function measureSortOutcome(page: Pick<Page, 'evaluate'>, click: () => Promise<void>, timeoutMs = 120_000): Promise<SortOutcome> {
  await page.evaluate(({ timeoutMs }) => {
    const host = window as SortWindow
    host.__viewerPerfSortWatch?.dispose()
    const header = () => [...document.querySelectorAll('[role=columnheader]')].find(h => /^Column G/.test(h.getAttribute('aria-label') ?? h.textContent ?? ''))
    const status = () => document.querySelector('[data-officeview-ui=viewer-status]')?.textContent ?? ''
    const before = header()?.getAttribute('aria-sort') ?? null, statusBefore = status()
    let pendingSeen = false, done = false, timer: number | undefined
    let resolve!: (result: Omit<SortOutcome, 'elapsedMs'>) => void
    const result = new Promise<Omit<SortOutcome, 'elapsedMs'>>(r => { resolve = r })
    const finish = (outcome: SortOutcome['outcome']) => {
      if (done) return
      done = true
      observer.disconnect()
      window.clearTimeout(timer)
      resolve({ outcome, status: status(), ariaSort: header()?.getAttribute('aria-sort') ?? null })
    }
    const check = () => {
      const text = status(), aria = header()?.getAttribute('aria-sort') ?? null
      if (header() && aria !== before) { finish('completed'); return }
      if (/Sorting column G…/.test(text)) { pendingSeen = true; return }
      // A fresh typed refusal also covers operations that end before React paints pending.
      const refusal = /Operation refused:|Operation cancelled\.|Source (?:data|range) is incomplete|Source or view changed;|Sort unavailable\./.test(text)
      if ((text !== statusBefore || pendingSeen) && !text.includes('View-only sort applied; formulas unchanged.') && (refusal || pendingSeen)) finish('refused')
    }
    const observer = new MutationObserver(check)
    observer.observe(document.body, { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: ['aria-sort'] })
    timer = window.setTimeout(() => finish('timeout'), timeoutMs)
    host.__viewerPerfSortWatch = { result, dispose: () => finish('timeout') }
  }, { timeoutMs })
  const started = Date.now()
  try {
    await click()
    const result = await page.evaluate(() => (window as SortWindow).__viewerPerfSortWatch!.result)
    return { ...result, elapsedMs: Date.now() - started }
  } finally {
    await page.evaluate(() => {
      const host = window as SortWindow
      host.__viewerPerfSortWatch?.dispose()
      delete host.__viewerPerfSortWatch
    })
  }
}

/** Keep failed operations visible in raw evidence and make the receipt inadmissible. */
export function recordSortOutcome(series: { operations: SortOutcome[]; endToEnd: number[] }, result: SortOutcome, errors: string[]): void {
  series.operations.push(result)
  if (result.outcome === 'completed') series.endToEnd.push(result.elapsedMs)
  else errors.push(`sort sample ${series.operations.length}: ${result.outcome} after ${result.elapsedMs} ms; ${result.status}`)
}
