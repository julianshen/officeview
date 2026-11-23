/**
 * TEST-ONLY options of the Playwright WS harness (server/test/app.ts), kept in a module so a server test can check them.
 *
 * `HARNESS_IP_RATES`: the per-IP session gate of the harness's stream servers. The product gate (§14, `IpRates`) is
 * 30 new sessions per minute per IP; the harness serves one browser running the whole stream project from 127.0.0.1,
 * one or more sessions per spec, so the suite crossed 30 in a minute at Stage 4b (the 31st handshake answered 429 and
 * `stream-text.spec.ts` "a11y=off" never painted). No browser spec exercises the per-IP gate (server/test/ip-rates.test.ts
 * does, at the product default). The concurrent-connect limit is left at the product's 8.
 */
export const HARNESS_IP_RATES = Object.freeze({ perMinute: 300 })
