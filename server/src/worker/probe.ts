/** Stage 1c Task 1 probe: @napi-rs/canvas paints inside a Bun Worker. */
import { createCanvas } from '@napi-rs/canvas'
import pkg from '@napi-rs/canvas/package.json'

declare const self: Worker
self.onmessage = () => {
  const c = createCanvas(4, 4), ctx = c.getContext('2d')
  ctx.fillStyle = '#ff0000'; ctx.fillRect(0, 0, 4, 4)
  postMessage({ rgba: Array.from(ctx.getImageData(0, 0, 4, 4).data), version: (pkg as { version: string }).version })
}
