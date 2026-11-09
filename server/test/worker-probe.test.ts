import { expect, it } from 'bun:test'
import { resolve } from 'node:path'

it('@napi-rs/canvas paints inside a Bun Worker', async () => {
  const w = new Worker(resolve(import.meta.dir, '../src/worker/probe.ts'))
  const reply = await new Promise<{ rgba: number[]; version: string }>((ok, fail) => {
    w.onmessage = (e) => ok(e.data); w.onerror = (e) => fail(e); w.postMessage(null)
  })
  w.terminate()
  expect(reply.rgba).toEqual(Array.from({ length: 16 }, () => [255, 0, 0, 255]).flat())
  expect(reply.version).toMatch(/^\d+\.\d+\.\d+/)
})
