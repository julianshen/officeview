// The documented high-level pako class, used only as a counter-oracle in tests
// (design note §2.4.1: its err === 0 answer must not be trusted).
declare module 'pako/lib/inflate.js' {
  export class Inflate {
    constructor(options?: { raw?: boolean })
    push(data: Uint8Array, mode?: boolean | number): boolean
    err: number
    ended: boolean
  }
}
