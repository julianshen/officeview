/**
 * Page-side result posting for the manual-device collector.
 *
 * The collector is mounted on the SAME ORIGIN as the page (a vite middleware
 * in vite.harness.config.ts), so the page posts to a fixed path and never to a
 * URL taken from the query string: an attacker-supplied `?collect=` link
 * cannot make the page send results to a foreign host. The per-session token
 * from the printed URL authorizes the POST.
 */
import type { HarnessResult } from './types'

export const COLLECT_PATH = '/__harness/collect'

export type PostOutcome = { status: 'no-collector' } | { status: 'posted' } | { status: 'rejected'; errors: string[] }

export async function postResult(
  result: HarnessResult,
  location: { href: string; search: string },
  fetchImpl: typeof fetch,
): Promise<PostOutcome> {
  const token = new URLSearchParams(location.search).get('token')
  if (!token) return { status: 'no-collector' }
  const response = await fetchImpl(new URL(COLLECT_PATH, location.href).toString(), {
    method: 'POST',
    headers: { 'content-type': 'application/json', 'x-harness-token': token },
    body: JSON.stringify(result),
  })
  if (response.ok) return { status: 'posted' }
  let errors = [`collector responded with HTTP ${response.status}`]
  try {
    const body = (await response.json()) as { errors?: unknown }
    if (Array.isArray(body.errors)) errors = body.errors.map(String)
  } catch {
    // keep the status-only message
  }
  return { status: 'rejected', errors }
}
