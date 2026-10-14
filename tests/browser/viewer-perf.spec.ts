/**
 * S10.2 performance page smoke (src/debug/viewer-perf.html under the harness
 * dev server). It guards the instrumentation the release bench relies on —
 * local-File open marks, open-loop scroll timelines, acknowledgment probes,
 * the worker ledger side channel, the importer pass, cross-origin isolated
 * memory measurement and teardown — with counts and states only. No wall-clock
 * assertion appears here: thresholds are applied by the release bench alone.
 */
import { expect, test, type Page } from '@playwright/test'
import { ensureFixtures, fixturePath } from './viewer-e2e'
import type { ViewerPerf } from '../../src/debug/viewer-perf'

type W = Window & { __viewerPerf: ViewerPerf }

async function boot(page: Page): Promise<void> {
  await page.goto('/src/debug/viewer-perf.html')
  await page.waitForFunction(() => (window as unknown as W).__viewerPerf?.booted === true)
}

test.describe('S10.2 viewer performance page', () => {
  test.beforeAll(async () => { await ensureFixtures() })

  test('worker-backed open, scroll timeline, acknowledgment, ledger, importer, memory and teardown', async ({ page, browserName }) => {
    await boot(page)
    await page.setInputFiles('#perf-file', fixturePath('f2-paged.xlsx'))
    await page.waitForFunction(() => (window as unknown as W).__viewerPerf.marks.completeAt !== null, undefined, { timeout: 90_000 })
    const marks = await page.evaluate(() => (window as unknown as W).__viewerPerf.marks)
    expect(marks.blobAt).not.toBeNull()
    expect(marks.firstUsableAt!).toBeGreaterThanOrEqual(marks.firstPresentAt!)
    expect(marks.usableStatus).toMatch(/Visible body/)

    const scroll = await page.evaluate(() => (window as unknown as W).__viewerPerf.scrollRun({ warmup: 5, frames: 30, dx: 0, dy: 60 }))
    expect(scroll.frameIntervals).toHaveLength(30)
    expect(scroll.scrollEvents).toBeGreaterThan(0)
    expect(scroll.inputToPaint.length).toBeGreaterThan(0)
    expect(scroll.timeline.presents.length).toBe(scroll.presents)

    await page.evaluate(() => (window as unknown as W).__viewerPerf.ackArm('dialog-open', 15_000))
    await page.getByRole('button', { name: /^Filter column D/ }).click()
    const ack = await page.evaluate(() => (window as unknown as W).__viewerPerf.ackWait())
    expect(ack.inputType).toMatch(/pointerdown|mousedown/)
    expect(ack.presentAt!).toBeGreaterThanOrEqual(ack.mutationAt!)
    await page.getByRole('dialog', { name: 'Column filter' }).getByRole('button', { name: 'Cancel' }).click()

    const stats = await page.evaluate(() => (window as unknown as W).__viewerPerf.workerStats())
    // Chromium keeps the paged document in its worker. Playwright WebKit may
    // refuse the worker route (capability, S10.3) and reopen the workbook on
    // the compatibility path: the worker is then already terminated.
    const workers = await page.evaluate(() => ({ created: (window as unknown as W).__viewerPerf.workersCreated, terminated: (window as unknown as W).__viewerPerf.workersTerminated }))
    if (browserName === 'chromium') expect(stats?.kind).toBe('stats')
    else if (stats === null) expect(workers.terminated).toBe(workers.created)
    if (stats?.kind === 'stats') {
      expect(stats.lifecycle.sessions).toBe(1)
      expect(stats.sessions[0].memory.limitBytes).toBe(512 * 1024 * 1024)
    }
    expect((await page.evaluate(() => (window as unknown as W).__viewerPerf.wire)).some(w => w.kind === 'open')).toBe(true)

    await page.setInputFiles('#perf-import-file', fixturePath('f1-small.xlsx'))
    const imported = await page.evaluate(() => (window as unknown as W).__viewerPerf.importRun('D'))
    if (browserName === 'chromium') expect(imported.kind).toBe('import')
    else if (imported.kind === 'error') expect(imported.message).toMatch(/\S/) // labelled, not hung
    if (imported.kind === 'import') {
      expect(imported.outcome).toBe('complete')
      expect(imported.events.map(e => e.phase)).toContain('done')
    }

    const memory = await page.evaluate(() => (window as unknown as W).__viewerPerf.measureMemory()) as { bytes?: number; unavailable?: string }
    // Installed Chrome measures; Playwright's bundled headless shell and WebKit
    // label the probe unavailable. Either way it is never a fabricated 0.
    if (memory.bytes !== undefined) expect(memory.bytes).toBeGreaterThan(0)
    else expect(memory.unavailable).toMatch(/\S/)
    if (browserName === 'chromium') expect(await page.evaluate(() => crossOriginIsolated)).toBe(true)

    await page.evaluate(() => (window as unknown as W).__viewerPerf.close())
    const after = await page.evaluate(() => ({ created: (window as unknown as W).__viewerPerf.workersCreated, terminated: (window as unknown as W).__viewerPerf.workersTerminated, canvas: (window as unknown as W).__viewerPerf.canvasBytes() }))
    expect(after.created).toBe(1)
    expect(after.terminated).toBe(after.created)
    expect(after.canvas).toEqual({ count: 0, bytes: 0 })
  })
})
